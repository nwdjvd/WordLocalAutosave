import time
import sys
import os
import json
import win32com.client
import win32com
from win32com.client import gencache, DispatchWithEvents
import pythoncom
import logging
from logging.handlers import RotatingFileHandler
from typing import Optional, Dict, Any, Iterator, Union, Callable

# Default configuration (will be overridden by config file if present)
DEFAULT_CONFIG = {
    "debounce_seconds": 10,       # Minimum time between saves for a document
    "polling_interval": 15,       # Seconds between backup polling checks
    "reconnect_threshold": 20,    # Number of attempts before reconnecting
    "main_loop_sleep": 0.5,       # Seconds to sleep in main loop
    "log_max_size": 5 * 1024 * 1024,  # 5 MB max log size
    "log_backup_count": 3,        # Number of log backups to keep
    "log_directory": "",          # Empty means current directory
}

def load_config(config_path: str = "autosave_config.json") -> Dict[str, Any]:
    """Load configuration from a JSON file, falling back to defaults.
    
    Args:
        config_path: Path to the configuration file
        
    Returns:
        Dict containing configuration values
    """
    config = DEFAULT_CONFIG.copy()
    
    # Try to load from config file
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                file_config = json.load(f)
                config.update(file_config)
            print(f"Loaded configuration from {config_path}")
        except Exception as e:
            print(f"Error loading configuration: {e}")
    
    # Check environment variables (override file config)
    for key in config:
        env_var = f"WORD_AUTOSAVE_{key.upper()}"
        if env_var in os.environ:
            try:
                # Convert string values to appropriate types
                if isinstance(config[key], bool):
                    config[key] = os.environ[env_var].lower() in ('true', 'yes', '1')
                elif isinstance(config[key], int):
                    config[key] = int(os.environ[env_var])
                elif isinstance(config[key], float):
                    config[key] = float(os.environ[env_var])
                else:
                    config[key] = os.environ[env_var]
            except ValueError:
                print(f"Invalid environment variable {env_var}")
    
    return config

# Setup logging
def setup_logging(config: Dict[str, Any]) -> logging.Logger:
    """Set up logging with rotation based on config.
    
    Args:
        config: Configuration dictionary
        
    Returns:
        Configured logger
    """
    log_dir = config["log_directory"]
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir, exist_ok=True)
    
    log_path = os.path.join(log_dir, "autosave.log") if log_dir else "autosave.log"
    
    logger = logging.getLogger("WordAutosave")
    logger.setLevel(logging.INFO)
    
    # Clear any existing handlers
    if logger.handlers:
        logger.handlers.clear()
    
    # File handler with rotation
    file_handler = RotatingFileHandler(
        log_path,
        maxBytes=config["log_max_size"],
        backupCount=config["log_backup_count"]
    )
    file_handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    ))
    logger.addHandler(file_handler)
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    ))
    logger.addHandler(console_handler)
    
    return logger

class WordWrapper:
    """Wrapper for Word COM application to facilitate testing."""
    
    def __init__(self, com_app: Any):
        """Initialize with Word application COM object.
        
        Args:
            com_app: Word Application COM object
        """
        self.app = com_app
    
    @property
    def name(self) -> str:
        """Get the name of the application."""
        return self.app.Name
    
    def document_count(self) -> int:
        """Get the number of open documents."""
        if not hasattr(self.app, 'Documents'):
            return 0
        return self.app.Documents.Count
    
    def get_document(self, index: int) -> Any:
        """Get a document by index.
        
        Args:
            index: 1-based index of the document
            
        Returns:
            Document COM object
        """
        return self.app.Documents.Item(index)
    
    def get_active_document(self) -> Any:
        """Get the active document.
        
        Returns:
            Active document COM object or None
        """
        if self.document_count() > 0:
            return self.app.ActiveDocument
        return None
    
    def iterate_documents(self) -> Iterator[Any]:
        """Safely iterate through all open documents.
        
        Yields:
            Document COM objects
        """
        count = self.document_count()
        for i in range(1, count + 1):
            try:
                yield self.get_document(i)
            except Exception:
                # Skip inaccessible documents
                continue

class WordAutosaveService:
    """Service that monitors and autosaves Word documents."""
    
    def __init__(self, config_path: str = "autosave_config.json"):
        """Initialize the autosave service.
        
        Args:
            config_path: Path to the configuration file
        """
        # Load configuration
        self.config = load_config(config_path)
        
        # Setup logging
        self.logger = setup_logging(self.config)
        
        # Internal state
        self.last_save_times = {}  # Tracks last save time per document
        self.word_app = None       # Word application COM object
        self.word = None           # Word wrapper
        self.connected = False     # Connection status to Word
        self.last_poll_time = 0    # Time of last polling check
        self.reconnect_attempts = 0  # Count of reconnection attempts
        self.running = True        # Flag for main loop control
        self.event_callbacks = {}  # Store event callbacks
    
    def on_window_selection_change(self, selection):
        """Triggered when text selection changes."""
        try:
            # Safely get the document from the selection
            try:
                doc = selection.Document
                if doc:
                    self.try_autosave(doc)
            except AttributeError:
                # Selection might not have a Document property
                self.logger.debug("Selection change event without valid document")
        except Exception:
            self.logger.exception("Error in OnWindowSelectionChange")
    
    def on_document_change(self):
        """Triggered when a document changes."""
        try:
            if not self.word:
                return
                
            doc = self.word.get_active_document()
            if doc:
                self.try_autosave(doc)
        except Exception:
            self.logger.exception("Error in OnDocumentChange")
    
    def on_document_open(self, doc):
        """Triggered when a document is opened."""
        try:
            self.try_autosave(doc)
            try:
                name = getattr(doc, 'Name', 'Unknown Document')
                self.logger.info(f"Document opened: {name}")
            except:
                self.logger.info("Document opened (name unknown)")
        except Exception:
            self.logger.exception("Error in OnDocumentOpen")
            
    def on_document_before_close(self, doc, cancel):
        """Triggered before a document is closed."""
        try:
            # Final save opportunity
            self.try_autosave(doc)
            
            # Remove from tracking
            try:
                path = getattr(doc, 'FullName', None)
                name = getattr(doc, 'Name', 'Unknown Document')
                
                if path and path in self.last_save_times:
                    del self.last_save_times[path]
                    self.logger.info(f"Document closed: {name}")
            except:
                self.logger.info("Document closed (name unknown)")
        except Exception:
            self.logger.exception("Error in OnDocumentBeforeClose")
    
    def on_document_before_save(self, doc, save_as_ui, cancel):
        """Triggered before a document is saved."""
        try:
            # Update the last save time when user manually saves
            try:
                path = getattr(doc, 'FullName', None)
                name = getattr(doc, 'Name', 'Unknown Document')
                
                if path:
                    self.last_save_times[path] = time.monotonic()
                    self.logger.debug(f"Manual save detected: {name}")
            except:
                self.logger.debug("Manual save detected (document name unknown)")
        except Exception:
            self.logger.exception("Error in OnDocumentBeforeSave")

    def try_autosave(self, doc: Any) -> bool:
        """Try to save the document if it's dirty and not recently saved.
        
        Args:
            doc: Word document COM object
            
        Returns:
            bool: True if document was saved, False otherwise
        """
        try:
            # Get document name and path for tracking
            try:
                name = getattr(doc, 'Name', 'Unknown Document')
                path = getattr(doc, 'FullName', f'Unknown_{id(doc)}')
            except:
                # Handle the case where doc might be a PyIDispatch without proper attribute access
                name = "Unknown Document"
                path = f"Unknown_{id(doc)}"
            
            # Check if document has unsaved changes - safely
            try:
                saved_state = doc.Saved
                if saved_state:  # If document is already saved
                    return False
            except:
                # Cannot determine saved state, assume we should try to save
                self.logger.warning(f"Couldn't check saved state for {name}, attempting save anyway")
            
            # Check debounce time
            now = time.monotonic()  # Use monotonic time for accurate intervals
            last_time = self.last_save_times.get(path, 0)
            if now - last_time < self.config["debounce_seconds"]:
                return False
            
            # Save document
            try:
                doc.Save()
                self.last_save_times[path] = now
                self.logger.info(f"Autosaved: {name}")
                return True
            except Exception as e:
                self.logger.exception(f"Error saving document '{name}': {str(e)}")
                return False
        except Exception:
            self.logger.exception(f"Unexpected error in try_autosave")
            return False

    def _process_events(self) -> None:
        """Process pending COM events."""
        pythoncom.PumpWaitingMessages()

    def _check_connection(self) -> bool:
        """Check if Word connection is still active.
        
        Returns:
            bool: True if connection is active, False otherwise
        """
        if not self.connected or not self.word:
            return False
            
        try:
            # Simple property access to test connection
            _ = self.word.name
            return True
        except Exception:
            self.logger.warning("Lost connection to Word")
            self.connected = False
            return False

    def _handle_reconnect(self) -> None:
        """Handle reconnection logic if connection is lost."""
        if self.connected:
            return
            
        self.reconnect_attempts += 1
        if self.reconnect_attempts > self.config["reconnect_threshold"]:
            if self.connect_to_word():
                self.logger.info("Reconnected to Word")
                self.reconnect_attempts = 0  # Reset only on successful reconnect

    def _poll_documents(self) -> None:
        """Poll all documents for changes and save if needed."""
        now = time.monotonic()
        if now - self.last_poll_time >= self.config["polling_interval"]:
            saved = self._check_all_documents()
            if saved > 0:
                self.logger.info(f"Poll check: saved {saved} documents")
            self.last_poll_time = now

    def _check_all_documents(self) -> int:
        """Check all open documents and save them if needed.
        
        Returns:
            int: Count of documents saved
        """
        if not self.word:
            return 0
            
        try:
            saved_count = 0
            for doc in self.word.iterate_documents():
                try:
                    if self.try_autosave(doc):
                        saved_count += 1
                except Exception:
                    self.logger.exception("Error checking document")
            
            return saved_count
        except Exception:
            self.logger.exception("Error iterating documents")
            return 0

    def _create_word_events_class(self):
        """Create a custom WordEvents class bound to this service instance.
        
        This is a simpler approach than trying to modify the event handler after creation.
        """
        # Define the event class with a closure to capture the service
        service = self
        
        class CustomWordEvents:
            """Event sink for Word application events."""
            
            def OnDocumentChange(self):
                service.on_document_change()
            
            def OnWindowSelectionChange(self, selection):
                service.on_window_selection_change(selection)
            
            def OnDocumentOpen(self, doc):
                service.on_document_open(doc)
            
            def OnDocumentBeforeClose(self, doc, cancel):
                service.on_document_before_close(doc, cancel)
            
            def OnDocumentBeforeSave(self, doc, saveAsUI, cancel):
                service.on_document_before_save(doc, saveAsUI, cancel)
        
        return CustomWordEvents

    def connect_to_word(self) -> bool:
        """Connect to Word application via COM.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            # Release any existing connections
            if self.word_app:
                try:
                    self.word_app = None
                    self.word = None
                except:
                    pass
            
            # Create a Word Application object with early binding for better type support
            word_app = gencache.EnsureDispatch("Word.Application")
            
            try:
                # Create a custom event class bound to this service
                CustomWordEvents = self._create_word_events_class()
                
                # Wrap the Word application with events
                self.word_app = DispatchWithEvents(word_app, CustomWordEvents)
                self.logger.info("Event handlers set up successfully")
                events_enabled = True
            except Exception:
                self.logger.exception("Event setup failed, falling back to polling only")
                self.word_app = word_app  # Use the app without events
                events_enabled = False
            
            # Create wrapper for easier testing and access
            self.word = WordWrapper(self.word_app)
            
            # If events are not enabled, use more frequent polling
            if not events_enabled:
                self.config["polling_interval"] = min(self.config["polling_interval"], 5)
                self.logger.info(f"Polling frequency set to every {self.config['polling_interval']} seconds")
            
            self.connected = True
            self.logger.info("Connected to Word successfully")
            return True
        except Exception:
            self.logger.exception("Failed to connect to Word")
            self.connected = False
            return False

    def run(self) -> None:
        """Main program loop."""
        self.logger.info("Word Local Autosave starting...")
        
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Initial connection
        self.connect_to_word()
        
        try:
            while self.running:
                try:
                    # Process COM events
                    self._process_events()
                    
                    # Check connection status
                    self._check_connection()
                    
                    # Handle reconnection if needed
                    self._handle_reconnect()
                    
                    # Poll documents periodically
                    self._poll_documents()
                    
                    # Sleep to reduce CPU usage
                    time.sleep(self.config["main_loop_sleep"])
                    
                except Exception:
                    self.logger.exception("Error in main loop")
                    time.sleep(1)  # Prevent tight error loop
        
        except KeyboardInterrupt:
            self.logger.info("Word Local Autosave stopping...")
        
        finally:
            # Proper cleanup
            self.shutdown()

    def stop(self) -> None:
        """Signal the service to stop."""
        self.running = False
        self.logger.info("Stop requested")

    def shutdown(self) -> None:
        """Clean shutdown of the service."""
        try:
            # Release COM objects
            if self.word_app:
                try:
                    self.word_app = None
                    self.word = None
                except:
                    pass
            
            # Explicitly uninitialize COM
            pythoncom.CoUninitialize()
            self.logger.info("Service shutdown complete")
        except Exception:
            self.logger.exception("Error during shutdown")

def main() -> None:
    """Entry point for the application."""
    # Print banner
    print(r"""
  __        __            _   _                 _   ____              
  \ \      / /__  _ __ __| | | |    ___   ___ __ | |/ ___|__ ___   _____
   \ \ /\ / / _ \| '__/ _` | | |   / _ \ / __/ _` | \___ / _` \ \ / / _ \
    \ V  V | (_) | | | (_| | | |__| (_) | (_| (_| | |__) (_| |\ V |  __/
     \_/\_/ \___/|_|  \__,_| |_____\___/ \___\__,_|_|____\__,_| \_/ \___|
                                                                        
                  Autosave service for Microsoft Word
    """)
    
    # Start the service
    service = WordAutosaveService()
    
    try:
        service.run()
    except KeyboardInterrupt:
        print("\nService stopped by user. Cleaning up...")
    except Exception as e:
        print(f"\nService crashed: {e}")
        print("See log file for details.")
    finally:
        # Ensure proper cleanup
        if service:
            service.shutdown()
        print("Service has stopped. Have a nice day!")

if __name__ == "__main__":
    main()
