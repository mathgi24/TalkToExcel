"""Hot-reload capability for operation config changes."""

import os
import time
import threading
from typing import Callable, Optional, Dict, Any
from pathlib import Path
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

from .template_loader import TemplateLoader
from .template_registry import TemplateRegistry
from .prompt_generator import PromptGenerator


class OperationConfigHandler(FileSystemEventHandler):
    """File system event handler for operation config changes."""
    
    def __init__(self, reload_callback: Callable[[], None]):
        """Initialize handler.
        
        Args:
            reload_callback: Function to call when config changes
        """
        self.reload_callback = reload_callback
        self.logger = logging.getLogger(__name__)
        self._last_reload = 0
        self._reload_delay = 1.0  # Minimum seconds between reloads
    
    def on_modified(self, event):
        """Handle file modification events."""
        if event.is_directory:
            return
        
        # Only react to YAML files
        if not event.src_path.endswith(('.yaml', '.yml')):
            return
        
        # Debounce rapid file changes
        current_time = time.time()
        if current_time - self._last_reload < self._reload_delay:
            return
        
        self._last_reload = current_time
        
        self.logger.info(f"Config file changed: {event.src_path}")
        
        try:
            self.reload_callback()
            self.logger.info("Configuration reloaded successfully")
        except Exception as e:
            self.logger.error(f"Failed to reload configuration: {e}")


class HotReloadManager:
    """Manages hot-reload functionality for operation configurations."""
    
    def __init__(self, template_loader: Optional[TemplateLoader] = None,
                 template_registry: Optional[TemplateRegistry] = None,
                 prompt_generator: Optional[PromptGenerator] = None):
        """Initialize hot-reload manager.
        
        Args:
            template_loader: TemplateLoader instance
            template_registry: TemplateRegistry instance  
            prompt_generator: PromptGenerator instance
        """
        self.template_loader = template_loader or TemplateLoader()
        self.template_registry = template_registry or TemplateRegistry(self.template_loader)
        self.prompt_generator = prompt_generator or PromptGenerator(
            self.template_loader, self.template_registry
        )
        
        self.logger = logging.getLogger(__name__)
        self._observer: Optional[Observer] = None
        self._is_watching = False
        self._reload_callbacks: Dict[str, Callable[[], None]] = {}
        
        # Set up default reload callback
        self.add_reload_callback('default', self._default_reload_callback)
    
    def _default_reload_callback(self) -> None:
        """Default callback that reloads all components."""
        self.template_loader.reload_operations()
        self.template_registry.reload_registry()
        # Note: PromptGenerator doesn't need explicit reload as it uses the updated loader/registry
    
    def add_reload_callback(self, name: str, callback: Callable[[], None]) -> None:
        """Add a callback to be executed on config reload.
        
        Args:
            name: Unique name for the callback
            callback: Function to call on reload
        """
        self._reload_callbacks[name] = callback
    
    def remove_reload_callback(self, name: str) -> bool:
        """Remove a reload callback.
        
        Args:
            name: Name of callback to remove
            
        Returns:
            True if callback was removed, False if not found
        """
        return self._reload_callbacks.pop(name, None) is not None
    
    def _execute_reload_callbacks(self) -> None:
        """Execute all registered reload callbacks."""
        for name, callback in self._reload_callbacks.items():
            try:
                callback()
                self.logger.debug(f"Executed reload callback: {name}")
            except Exception as e:
                self.logger.error(f"Error in reload callback '{name}': {e}")
    
    def start_watching(self, watch_directory: Optional[str] = None) -> bool:
        """Start watching for configuration file changes.
        
        Args:
            watch_directory: Directory to watch. Defaults to templates directory.
            
        Returns:
            True if watching started successfully
        """
        if self._is_watching:
            self.logger.warning("Hot-reload is already active")
            return True
        
        if watch_directory is None:
            watch_directory = str(self.template_loader.templates_dir)
        
        watch_path = Path(watch_directory)
        if not watch_path.exists():
            self.logger.error(f"Watch directory does not exist: {watch_path}")
            return False
        
        try:
            self._observer = Observer()
            event_handler = OperationConfigHandler(self._execute_reload_callbacks)
            self._observer.schedule(event_handler, str(watch_path), recursive=True)
            self._observer.start()
            
            self._is_watching = True
            self.logger.info(f"Started watching for config changes in: {watch_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to start file watching: {e}")
            return False
    
    def stop_watching(self) -> bool:
        """Stop watching for configuration file changes.
        
        Returns:
            True if watching stopped successfully
        """
        if not self._is_watching or self._observer is None:
            self.logger.warning("Hot-reload is not active")
            return True
        
        try:
            self._observer.stop()
            self._observer.join(timeout=5.0)
            self._observer = None
            self._is_watching = False
            
            self.logger.info("Stopped watching for config changes")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to stop file watching: {e}")
            return False
    
    def is_watching(self) -> bool:
        """Check if hot-reload is currently active.
        
        Returns:
            True if watching for changes
        """
        return self._is_watching
    
    def manual_reload(self) -> bool:
        """Manually trigger a configuration reload.
        
        Returns:
            True if reload was successful
        """
        try:
            self._execute_reload_callbacks()
            self.logger.info("Manual configuration reload completed")
            return True
        except Exception as e:
            self.logger.error(f"Manual reload failed: {e}")
            return False
    
    def get_reload_status(self) -> Dict[str, Any]:
        """Get current hot-reload status information.
        
        Returns:
            Dictionary with status information
        """
        return {
            'is_watching': self._is_watching,
            'watch_directory': str(self.template_loader.templates_dir),
            'callbacks_registered': len(self._reload_callbacks),
            'callback_names': list(self._reload_callbacks.keys()),
            'last_modified': self.template_loader._last_modified.isoformat() if self.template_loader._last_modified else None
        }
    
    def check_for_changes(self) -> bool:
        """Check if configuration files have been modified.
        
        Returns:
            True if files have been modified since last load
        """
        return self.template_loader.is_file_modified()
    
    def __enter__(self):
        """Context manager entry."""
        self.start_watching()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.stop_watching()


class PeriodicReloadChecker:
    """Periodically checks for configuration changes without file system watching."""
    
    def __init__(self, hot_reload_manager: HotReloadManager, 
                 check_interval: float = 5.0):
        """Initialize periodic checker.
        
        Args:
            hot_reload_manager: HotReloadManager instance
            check_interval: Seconds between checks
        """
        self.hot_reload_manager = hot_reload_manager
        self.check_interval = check_interval
        self.logger = logging.getLogger(__name__)
        
        self._timer: Optional[threading.Timer] = None
        self._is_running = False
    
    def _check_and_reload(self) -> None:
        """Check for changes and reload if necessary."""
        try:
            if self.hot_reload_manager.check_for_changes():
                self.logger.info("Configuration changes detected, reloading...")
                self.hot_reload_manager.manual_reload()
        except Exception as e:
            self.logger.error(f"Error during periodic reload check: {e}")
        finally:
            # Schedule next check
            if self._is_running:
                self._schedule_next_check()
    
    def _schedule_next_check(self) -> None:
        """Schedule the next periodic check."""
        self._timer = threading.Timer(self.check_interval, self._check_and_reload)
        self._timer.daemon = True
        self._timer.start()
    
    def start(self) -> None:
        """Start periodic checking."""
        if self._is_running:
            self.logger.warning("Periodic checker is already running")
            return
        
        self._is_running = True
        self._schedule_next_check()
        self.logger.info(f"Started periodic config checking (interval: {self.check_interval}s)")
    
    def stop(self) -> None:
        """Stop periodic checking."""
        if not self._is_running:
            self.logger.warning("Periodic checker is not running")
            return
        
        self._is_running = False
        if self._timer:
            self._timer.cancel()
            self._timer = None
        
        self.logger.info("Stopped periodic config checking")
    
    def is_running(self) -> bool:
        """Check if periodic checking is active.
        
        Returns:
            True if checking is active
        """
        return self._is_running