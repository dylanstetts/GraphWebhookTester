#!/usr/bin/env python3
"""
Microsoft Graph Security Webhook Tester
A GUI application to test Microsoft Graph change notifications for security-related user actions.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import requests
import json
import logging
import os
from datetime import datetime, timedelta
import threading
import time
import webbrowser
from typing import Dict, Any, Optional
import urllib.parse
import pygame
from pathlib import Path

# Import delta tracker for detailed change analysis
try:
    from enhanced_change_tracker import EnhancedChangeTracker
except ImportError:
    print("Warning: enhanced_change_tracker.py not found. Enhanced analysis will be limited.")
    EnhancedChangeTracker = None

# Import MSAL for authentication
try:
    import msal
except ImportError:
    print("MSAL not installed. Please install it with: pip install msal")
    exit(1)

class HTTPLogger:
    """Custom logger for HTTP requests and responses"""
    
    def __init__(self, log_file: str = "logs/graph_api_requests.log"):
        # Make log file path relative to script directory
        if not os.path.isabs(log_file):
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self.log_file = os.path.join(script_dir, log_file)
        else:
            self.log_file = log_file
        
        # Create logs directory if it doesn't exist
        log_dir = os.path.dirname(self.log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        self.logger = logging.getLogger("GraphAPILogger")
        self.logger.setLevel(logging.INFO)
        
        # Create file handler
        handler = logging.FileHandler(self.log_file, encoding='utf-8')
        handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)
        
        # Add handler to logger
        if not self.logger.handlers:
            self.logger.addHandler(handler)
    
    def log_request(self, method: str, url: str, headers: Dict, body: str = None):
        """Log HTTP request"""
        self.logger.info(f"=== REQUEST ===")
        self.logger.info(f"Method: {method}")
        self.logger.info(f"URL: {url}")
        self.logger.info(f"Headers: {json.dumps(dict(headers), indent=2)}")
        if body:
            self.logger.info(f"Body: {body}")
        self.logger.info(f"==================")
    
    def log_response(self, status_code: int, headers: Dict, body: str):
        """Log HTTP response"""
        self.logger.info(f"=== RESPONSE ===")
        self.logger.info(f"Status Code: {status_code}")
        self.logger.info(f"Headers: {json.dumps(dict(headers), indent=2)}")
        self.logger.info(f"Body: {body}")
        self.logger.info(f"===================")

class GraphAuthenticator:
    """Handles Microsoft Graph authentication using MSAL"""
    
    def __init__(self, client_id: str, client_secret: str = None, tenant_id: str = "common"):
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scopes = [
            "https://graph.microsoft.com/Subscription.Read.All",
            "https://graph.microsoft.com/Subscription.ReadWrite.All",
            "https://graph.microsoft.com/Files.ReadWrite.All",
            "https://graph.microsoft.com/Sites.ReadWrite.All"
        ]
        self.access_token = None
        
        # Initialize MSAL app
        if client_secret:
            # Confidential client app
            self.app = msal.ConfidentialClientApplication(
                client_id=client_id,
                client_credential=client_secret,
                authority=self.authority
            )
        else:
            # Public client app
            self.app = msal.PublicClientApplication(
                client_id=client_id,
                authority=self.authority
            )
    
    def authenticate_interactive(self) -> bool:
        """Authenticate using interactive flow"""
        try:
            # Try to get token from cache first
            accounts = self.app.get_accounts()
            if accounts:
                result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
                if result and "access_token" in result:
                    self.access_token = result["access_token"]
                    return True
            
            # Interactive authentication
            result = self.app.acquire_token_interactive(
                scopes=self.scopes,
                prompt="select_account"
            )
            
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                return True
            else:
                error = result.get("error", "Unknown error")
                error_desc = result.get("error_description", "No description")
                raise Exception(f"Authentication failed: {error} - {error_desc}")
                
        except Exception as e:
            raise Exception(f"Authentication error: {str(e)}")
    
    def authenticate_client_credentials(self) -> bool:
        """Authenticate using client credentials flow (app-only)"""
        if not self.client_secret:
            raise Exception("Client secret required for client credentials flow")
        
        try:
            result = self.app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                return True
            else:
                error = result.get("error", "Unknown error")
                error_desc = result.get("error_description", "No description")
                raise Exception(f"Authentication failed: {error} - {error_desc}")
                
        except Exception as e:
            raise Exception(f"Authentication error: {str(e)}")

class GraphSubscriptionManager:
    """Manages Microsoft Graph subscriptions"""
    
    def __init__(self, authenticator: GraphAuthenticator, http_logger: HTTPLogger):
        self.authenticator = authenticator
        self.http_logger = http_logger
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"
    
    def create_subscription(self, 
                          resource: str, 
                          change_type: str, 
                          notification_url: str,
                          expiration_hours: int = 24,
                          include_security_webhooks: bool = True) -> Dict[str, Any]:
        """Create a Microsoft Graph subscription"""
        
        if not self.authenticator.access_token:
            raise Exception("Not authenticated. Please authenticate first.")
        
        # Calculate expiration time
        expiration = datetime.utcnow() + timedelta(hours=expiration_hours)
        expiration_str = expiration.strftime("%Y-%m-%dT%H:%M:%S.0000000Z")
        
        # Prepare subscription payload
        subscription_data = {
            "changeType": change_type,
            "notificationUrl": notification_url,
            "resource": resource,
            "expirationDateTime": expiration_str,
            "clientState": "webhook-test-" + str(int(time.time()))
        }
        
        headers = {
            "Authorization": f"Bearer {self.authenticator.access_token}",
            "Content-Type": "application/json"
        }
        
        # Add security webhooks header if requested
        if include_security_webhooks:
            headers["Prefer"] = "includesecuritywebhooks"
        
        url = f"{self.graph_endpoint}/subscriptions"
        body = json.dumps(subscription_data, indent=2)
        
        # Log the request
        self.http_logger.log_request("POST", url, headers, body)
        
        try:
            response = requests.post(url, headers=headers, data=body)
            
            # Log the response
            response_body = response.text
            self.http_logger.log_response(response.status_code, response.headers, response_body)
            
            if response.status_code == 201:
                return {
                    "success": True,
                    "data": response.json(),
                    "status_code": response.status_code,
                    "message": "Subscription created successfully!"
                }
            else:
                error_data = response.json() if response.text else {}
                return {
                    "success": False,
                    "error": error_data,
                    "status_code": response.status_code,
                    "message": f"Failed to create subscription. Status: {response.status_code}"
                }
                
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": str(e),
                "status_code": None,
                "message": f"Request failed: {str(e)}"
            }
        except json.JSONDecodeError as e:
            return {
                "success": False,
                "error": str(e),
                "status_code": response.status_code,
                "message": f"Invalid JSON response: {str(e)}"
            }
    
    def list_subscriptions(self) -> Dict[str, Any]:
        """List all current subscriptions"""
        
        if not self.authenticator.access_token:
            raise Exception("Not authenticated. Please authenticate first.")
        
        headers = {
            "Authorization": f"Bearer {self.authenticator.access_token}",
            "Content-Type": "application/json"
        }
        
        url = f"{self.graph_endpoint}/subscriptions"
        
        # Log the request
        self.http_logger.log_request("GET", url, headers)
        
        try:
            response = requests.get(url, headers=headers)
            
            # Log the response
            response_body = response.text
            self.http_logger.log_response(response.status_code, response.headers, response_body)
            
            if response.status_code == 200:
                return {
                    "success": True,
                    "data": response.json(),
                    "status_code": response.status_code,
                    "message": "Subscriptions retrieved successfully!"
                }
            else:
                error_data = response.json() if response.text else {}
                return {
                    "success": False,
                    "error": error_data,
                    "status_code": response.status_code,
                    "message": f"Failed to retrieve subscriptions. Status: {response.status_code}"
                }
                
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": str(e),
                "status_code": None,
                "message": f"Request failed: {str(e)}"
            }
        except json.JSONDecodeError as e:
            return {
                "success": False,
                "error": str(e),
                "status_code": response.status_code,
                "message": f"Invalid JSON response: {str(e)}"
            }
    
    def delete_subscription(self, subscription_id: str) -> Dict[str, Any]:
        """Delete a Microsoft Graph subscription"""
        
        if not self.authenticator.access_token:
            raise Exception("Not authenticated. Please authenticate first.")
        
        headers = {
            "Authorization": f"Bearer {self.authenticator.access_token}",
            "Content-Type": "application/json"
        }
        
        url = f"{self.graph_endpoint}/subscriptions/{subscription_id}"
        
        # Log the request
        self.http_logger.log_request("DELETE", url, headers)
        
        try:
            response = requests.delete(url, headers=headers)
            
            # Log the response
            response_body = response.text if response.text else ""
            self.http_logger.log_response(response.status_code, response.headers, response_body)
            
            if response.status_code == 204:
                return {
                    "success": True,
                    "data": None,
                    "status_code": response.status_code,
                    "message": "Subscription deleted successfully!"
                }
            else:
                error_data = response.json() if response.text else {}
                return {
                    "success": False,
                    "error": error_data,
                    "status_code": response.status_code,
                    "message": f"Failed to delete subscription. Status: {response.status_code}"
                }
                
        except requests.exceptions.RequestException as e:
            return {
                "success": False,
                "error": str(e),
                "status_code": None,
                "message": f"Request failed: {str(e)}"
            }
        except json.JSONDecodeError as e:
            return {
                "success": False,
                "error": str(e),
                "status_code": response.status_code,
                "message": f"Invalid JSON response: {str(e)}"
            }

class SoundManager:
    """Manages sound notifications"""
    
    def __init__(self):
        pygame.mixer.init()
        self.sounds = {}
        self._load_default_sounds()
    
    def _load_default_sounds(self):
        """Load default system sounds"""
        try:
            # Create simple tone sounds using pygame
            pass
        except Exception as e:
            print(f"Could not load sounds: {e}")
    
    def play_success_sound(self):
        """Play success notification sound"""
        try:
            # Create a simple success tone
            import pygame.sndarray
            import numpy as np
            
            sample_rate = 22050
            duration = 0.5
            frequency = 800
            
            frames = int(duration * sample_rate)
            arr = np.zeros((frames, 2))
            
            for i in range(frames):
                time_val = float(i) / sample_rate
                wave = np.sin(frequency * 2 * np.pi * time_val)
                arr[i] = [wave, wave]
            
            arr = (arr * 32767).astype(np.int16)
            sound = pygame.sndarray.make_sound(arr)
            sound.play()
            
        except Exception as e:
            # Fallback to system beep
            print('\a')  # System beep
    
    def play_error_sound(self):
        """Play error notification sound"""
        try:
            # Create a simple error tone
            import pygame.sndarray
            import numpy as np
            
            sample_rate = 22050
            duration = 0.3
            frequency = 400
            
            frames = int(duration * sample_rate)
            arr = np.zeros((frames, 2))
            
            for i in range(frames):
                time_val = float(i) / sample_rate
                wave = np.sin(frequency * 2 * np.pi * time_val)
                arr[i] = [wave, wave]
            
            arr = (arr * 32767).astype(np.int16)
            sound = pygame.sndarray.make_sound(arr)
            sound.play()
            
        except Exception as e:
            # Fallback to system beep
            print('\a\a')  # Double system beep

class GraphWebhookTesterGUI:
    """Main GUI application for testing Microsoft Graph webhooks"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Microsoft Graph Security Webhook Tester")
        self.root.geometry("900x700")
        
        # Initialize components
        self.http_logger = HTTPLogger()
        self.sound_manager = SoundManager()
        self.authenticator = None
        self.subscription_manager = None
        
        # Create GUI
        self._create_gui()
        
        # Load configuration if exists
        self._load_config()
    
    def _create_gui(self):
        """Create the GUI interface"""
        
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Authentication tab
        auth_frame = ttk.Frame(notebook)
        notebook.add(auth_frame, text="Authentication")
        self._create_auth_tab(auth_frame)
        
        # Subscription tab
        sub_frame = ttk.Frame(notebook)
        notebook.add(sub_frame, text="Create Subscription")
        self._create_subscription_tab(sub_frame)
        
        # Monitor tab
        monitor_frame = ttk.Frame(notebook)
        notebook.add(monitor_frame, text="Monitor Subscriptions")
        self._create_monitor_tab(monitor_frame)
        
        # Logs tab
        logs_frame = ttk.Frame(notebook)
        notebook.add(logs_frame, text="API Logs")
        self._create_logs_tab(logs_frame)
        
        # Delta query tab for change details
        delta_frame = ttk.Frame(notebook)
        notebook.add(delta_frame, text="Change Details")
        self._create_delta_tab(delta_frame)
    
    def _create_auth_tab(self, parent):
        """Create authentication tab"""
        
        # Title
        title_label = ttk.Label(parent, text="Microsoft Graph Authentication", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Configuration frame
        config_frame = ttk.LabelFrame(parent, text="App Registration Configuration")
        config_frame.pack(fill="x", padx=10, pady=5)
        
        # Client ID
        ttk.Label(config_frame, text="Client ID:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.client_id_var = tk.StringVar()
        client_id_entry = ttk.Entry(config_frame, textvariable=self.client_id_var, width=50)
        client_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Client Secret (optional)
        ttk.Label(config_frame, text="Client Secret (Optional):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.client_secret_var = tk.StringVar()
        client_secret_entry = ttk.Entry(config_frame, textvariable=self.client_secret_var, width=50, show="*")
        client_secret_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Tenant ID
        ttk.Label(config_frame, text="Tenant ID:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.tenant_id_var = tk.StringVar(value="common")
        tenant_id_entry = ttk.Entry(config_frame, textvariable=self.tenant_id_var, width=50)
        tenant_id_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # Authentication type
        auth_type_frame = ttk.LabelFrame(parent, text="Authentication Type")
        auth_type_frame.pack(fill="x", padx=10, pady=5)
        
        self.auth_type_var = tk.StringVar(value="interactive")
        ttk.Radiobutton(auth_type_frame, text="Interactive (User)", variable=self.auth_type_var, value="interactive").pack(anchor="w", padx=5, pady=2)
        ttk.Radiobutton(auth_type_frame, text="App-only (Client Credentials)", variable=self.auth_type_var, value="client_credentials").pack(anchor="w", padx=5, pady=2)
        
        # Buttons frame
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(buttons_frame, text="Save Configuration", command=self._save_config).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Load Configuration", command=self._load_config_dialog).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Authenticate", command=self._authenticate).pack(side="left", padx=5)
        
        # Status frame
        status_frame = ttk.LabelFrame(parent, text="Authentication Status")
        status_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.auth_status_text = scrolledtext.ScrolledText(status_frame, height=10, state="disabled")
        self.auth_status_text.pack(fill="both", expand=True, padx=5, pady=5)
    
    def _create_subscription_tab(self, parent):
        """Create subscription creation tab"""
        
        # Title
        title_label = ttk.Label(parent, text="Create Microsoft Graph Subscription", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Subscription configuration
        config_frame = ttk.LabelFrame(parent, text="Subscription Configuration")
        config_frame.pack(fill="x", padx=10, pady=5)
        
        # Resource
        ttk.Label(config_frame, text="Resource:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        resource_frame = ttk.Frame(config_frame)
        resource_frame.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        self.resource_var = tk.StringVar(value="/me/drive/root")
        resource_combo = ttk.Combobox(resource_frame, textvariable=self.resource_var, width=57,
                                    values=[
                                        "/me/drive/root",
                                        "/me/drive/items/{item-id}",
                                        "/sites/{site-id}/drive/root",
                                        "/groups/{group-id}/drive/root",
                                        "/users/{user-id}/drive/root"
                                    ])
        resource_combo.pack(side="left")
        
        ttk.Button(resource_frame, text="Help", width=6, 
                  command=self._show_resource_help).pack(side="left", padx=(5, 0))
        
        # Change Type
        ttk.Label(config_frame, text="Change Type:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.change_type_var = tk.StringVar(value="updated")
        change_type_combo = ttk.Combobox(config_frame, textvariable=self.change_type_var, 
                                       values=["updated", "deleted", "updated,deleted"])
        change_type_combo.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Notification URL
        ttk.Label(config_frame, text="Notification URL:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        notification_frame = ttk.Frame(config_frame)
        notification_frame.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        self.notification_url_var = tk.StringVar(value="https://webhook.site/unique-id")
        notification_url_combo = ttk.Combobox(notification_frame, textvariable=self.notification_url_var, width=57,
                                            values=[
                                                "http://localhost:8000",
                                                "https://webhook.site/unique-id",
                                                "https://your-app.azurewebsites.net/webhook",
                                                "https://your-domain.com/webhook"
                                            ])
        notification_url_combo.pack(side="left")
        
        ttk.Button(notification_frame, text="Help", width=6, 
                  command=self._show_webhook_help).pack(side="left", padx=(5, 0))
        
        # Expiration hours
        ttk.Label(config_frame, text="Expiration (hours):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.expiration_var = tk.StringVar(value="24")
        expiration_entry = ttk.Entry(config_frame, textvariable=self.expiration_var, width=10)
        expiration_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # Security webhooks option
        security_frame = ttk.LabelFrame(parent, text="Security Options")
        security_frame.pack(fill="x", padx=10, pady=5)
        
        self.include_security_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(security_frame, text="Include Security Webhooks (Prefer: includesecuritywebhooks)", 
                       variable=self.include_security_var).pack(anchor="w", padx=5, pady=5)
        
        # Configuration management buttons
        config_buttons_frame = ttk.Frame(parent)
        config_buttons_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(config_buttons_frame, text="Load Defaults", command=self._load_defaults).pack(side="left", padx=5)
        ttk.Button(config_buttons_frame, text="Save as Defaults", command=self._save_as_defaults).pack(side="left", padx=5)
        ttk.Button(config_buttons_frame, text="Reset Fields", command=self._reset_subscription_fields).pack(side="left", padx=5)
        
        # Create button
        create_button = ttk.Button(parent, text="Create Subscription", command=self._create_subscription)
        create_button.pack(pady=10)
        
        # Response frame
        response_frame = ttk.LabelFrame(parent, text="Response")
        response_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.response_text = scrolledtext.ScrolledText(response_frame, height=15, state="disabled")
        self.response_text.pack(fill="both", expand=True, padx=5, pady=5)
    
    def _create_monitor_tab(self, parent):
        """Create subscription monitoring tab"""
        
        # Title
        title_label = ttk.Label(parent, text="Monitor Active Subscriptions", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Buttons
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Refresh Subscriptions", command=self._refresh_subscriptions).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Delete Selected", command=self._delete_selected_subscription).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Delete All", command=self._delete_all_subscriptions).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Open Log File", command=self._open_log_file).pack(side="left", padx=5)
        
        # Subscription selection frame
        selection_frame = ttk.LabelFrame(parent, text="Subscription Selection")
        selection_frame.pack(fill="x", padx=10, pady=5)
        
        # Subscription dropdown
        subscription_select_frame = ttk.Frame(selection_frame)
        subscription_select_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(subscription_select_frame, text="Select Subscription:").pack(side="left", padx=(0, 5))
        self.subscription_id_var = tk.StringVar()
        self.subscription_combo = ttk.Combobox(subscription_select_frame, textvariable=self.subscription_id_var, width=50, state="readonly")
        self.subscription_combo.pack(side="left", padx=5)
        
        ttk.Button(subscription_select_frame, text="Refresh List", command=self._refresh_subscription_dropdown).pack(side="left", padx=5)
        
        # Subscriptions list
        list_frame = ttk.LabelFrame(parent, text="Active Subscriptions")
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.subscriptions_text = scrolledtext.ScrolledText(list_frame, height=15, state="disabled")
        self.subscriptions_text.pack(fill="both", expand=True, padx=5, pady=5)
    
    def _create_logs_tab(self, parent):
        """Create API logs tab"""
        
        # Title
        title_label = ttk.Label(parent, text="Microsoft Graph API Request/Response Logs", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Buttons
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Refresh Logs", command=self._refresh_logs).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Clear Logs", command=self._clear_logs).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Export Logs", command=self._export_logs).pack(side="left", padx=5)
        
        # Logs display
        logs_frame = ttk.LabelFrame(parent, text="API Logs")
        logs_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.logs_text = scrolledtext.ScrolledText(logs_frame, height=25, state="disabled")
        self.logs_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Auto-refresh logs
        self._refresh_logs()
    
    def _create_delta_tab(self, parent):
        """Create delta query/change details tab"""
        
        # Title
        title_label = ttk.Label(parent, text="Detailed Change Analysis", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)
        
        # Instructions
        instructions = ttk.Label(parent, 
                               text="This tab shows detailed analysis of what specifically changed, obtained via Microsoft Graph Delta Query API.\n"
                                    "When webhook notifications are received, the system automatically queries for detailed changes.",
                               wraplength=700, justify="left")
        instructions.pack(padx=10, pady=5)
        
        # Control buttons
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(buttons_frame, text="Analyze Latest Webhook", command=self._analyze_latest_webhook).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Refresh Changes", command=self._refresh_changes).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Clear Analysis", command=self._clear_change_analysis).pack(side="left", padx=5)
        
        # File selection frame
        file_frame = ttk.LabelFrame(parent, text="Analyze Specific Webhook File")
        file_frame.pack(fill="x", padx=10, pady=5)
        
        file_select_frame = ttk.Frame(file_frame)
        file_select_frame.pack(fill="x", padx=5, pady=5)
        
        self.webhook_file_var = tk.StringVar()
        webhook_file_combo = ttk.Combobox(file_select_frame, textvariable=self.webhook_file_var, width=50)
        webhook_file_combo.pack(side="left", padx=(0, 5))
        
        ttk.Button(file_select_frame, text="Browse", command=self._browse_webhook_file).pack(side="left", padx=5)
        ttk.Button(file_select_frame, text="Analyze Selected", command=self._analyze_selected_webhook).pack(side="left", padx=5)
        
        # Update webhook file list
        self._update_webhook_file_list(webhook_file_combo)
        
        # Change details display
        details_frame = ttk.LabelFrame(parent, text="Change Details")
        details_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.change_details_text = scrolledtext.ScrolledText(details_frame, height=20, state="disabled")
        self.change_details_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Auto-refresh changes
        self._refresh_changes()
    
    def _authenticate(self):
        """Authenticate with Microsoft Graph"""
        
        def auth_worker():
            try:
                self._update_auth_status("Starting authentication...")
                
                client_id = self.client_id_var.get().strip()
                client_secret = self.client_secret_var.get().strip() or None
                tenant_id = self.tenant_id_var.get().strip() or "common"
                auth_type = self.auth_type_var.get()
                
                if not client_id:
                    raise Exception("Client ID is required")
                
                # Initialize authenticator
                self.authenticator = GraphAuthenticator(client_id, client_secret, tenant_id)
                
                # Authenticate based on type
                if auth_type == "interactive":
                    self._update_auth_status("Opening browser for interactive authentication...")
                    success = self.authenticator.authenticate_interactive()
                else:
                    if not client_secret:
                        raise Exception("Client secret is required for app-only authentication")
                    self._update_auth_status("Authenticating with client credentials...")
                    success = self.authenticator.authenticate_client_credentials()
                
                if success:
                    # Initialize subscription manager
                    self.subscription_manager = GraphSubscriptionManager(self.authenticator, self.http_logger)
                    
                    self._update_auth_status("Authentication successful!")
                    self.sound_manager.play_success_sound()
                else:
                    self._update_auth_status("Authentication failed!")
                    self.sound_manager.play_error_sound()
                    
            except Exception as e:
                self._update_auth_status(f"Authentication error: {str(e)}")
                self.sound_manager.play_error_sound()
        
        # Run authentication in background thread
        threading.Thread(target=auth_worker, daemon=True).start()
    
    def _create_subscription(self):
        """Create a Microsoft Graph subscription"""
        
        def create_worker():
            try:
                if not self.subscription_manager:
                    raise Exception("Please authenticate first")
                
                self._update_response("Creating subscription...")
                
                resource = self.resource_var.get().strip()
                change_type = self.change_type_var.get().strip()
                notification_url = self.notification_url_var.get().strip()
                expiration_hours = int(self.expiration_var.get().strip())
                include_security = self.include_security_var.get()
                
                if not all([resource, change_type, notification_url]):
                    raise Exception("Resource, change type, and notification URL are required")
                
                # Create subscription
                result = self.subscription_manager.create_subscription(
                    resource=resource,
                    change_type=change_type,
                    notification_url=notification_url,
                    expiration_hours=expiration_hours,
                    include_security_webhooks=include_security
                )
                
                # Format response
                response_text = f"Status: {result['message']}\n"
                response_text += f"Success: {result['success']}\n"
                response_text += f"Status Code: {result['status_code']}\n\n"
                
                if result['success']:
                    response_text += "Subscription Details:\n"
                    response_text += json.dumps(result['data'], indent=2)
                    self.sound_manager.play_success_sound()
                else:
                    response_text += "Error Details:\n"
                    response_text += json.dumps(result.get('error', {}), indent=2)
                    self.sound_manager.play_error_sound()
                
                self._update_response(response_text)
                
            except Exception as e:
                error_text = f"Error creating subscription: {str(e)}"
                self._update_response(error_text)
                self.sound_manager.play_error_sound()
        
        # Run in background thread
        threading.Thread(target=create_worker, daemon=True).start()
    
    def _refresh_subscriptions(self):
        """Refresh the list of active subscriptions"""
        
        def refresh_worker():
            try:
                if not self.subscription_manager:
                    raise Exception("Please authenticate first")
                
                self._update_subscriptions("Refreshing subscriptions...")
                
                result = self.subscription_manager.list_subscriptions()
                
                if result['success']:
                    subscriptions = result['data'].get('value', [])
                    
                    if not subscriptions:
                        text = "No active subscriptions found."
                    else:
                        text = f"Found {len(subscriptions)} active subscription(s):\n\n"
                        for i, sub in enumerate(subscriptions, 1):
                            text += f"Subscription {i}:\n"
                            text += f"  ID: {sub.get('id', 'N/A')}\n"
                            text += f"  Resource: {sub.get('resource', 'N/A')}\n"
                            text += f"  Change Type: {sub.get('changeType', 'N/A')}\n"
                            text += f"  Notification URL: {sub.get('notificationUrl', 'N/A')}\n"
                            text += f"  Expiration: {sub.get('expirationDateTime', 'N/A')}\n"
                            text += f"  Client State: {sub.get('clientState', 'N/A')}\n"
                            text += "-" * 60 + "\n"
                    
                    self._update_subscriptions(text)
                else:
                    error_text = f"Error retrieving subscriptions: {result['message']}\n"
                    error_text += json.dumps(result.get('error', {}), indent=2)
                    self._update_subscriptions(error_text)
                    
            except Exception as e:
                error_text = f"Error refreshing subscriptions: {str(e)}"
                self._update_subscriptions(error_text)
        
        # Run in background thread
        threading.Thread(target=refresh_worker, daemon=True).start()
    
    def _refresh_logs(self):
        """Refresh the API logs display"""
        try:
            if os.path.exists(self.http_logger.log_file):
                with open(self.http_logger.log_file, 'r', encoding='utf-8') as f:
                    log_content = f.read()
                    
                self.logs_text.config(state="normal")
                self.logs_text.delete(1.0, tk.END)
                self.logs_text.insert(tk.END, log_content)
                self.logs_text.config(state="disabled")
                
                # Scroll to bottom
                self.logs_text.see(tk.END)
            else:
                self.logs_text.config(state="normal")
                self.logs_text.delete(1.0, tk.END)
                self.logs_text.insert(tk.END, "No log file found. Make API requests to see logs here.")
                self.logs_text.config(state="disabled")
                
        except Exception as e:
            self.logs_text.config(state="normal")
            self.logs_text.delete(1.0, tk.END)
            self.logs_text.insert(tk.END, f"Error reading log file: {str(e)}")
            self.logs_text.config(state="disabled")
    
    def _clear_logs(self):
        """Clear the API logs"""
        try:
            if os.path.exists(self.http_logger.log_file):
                open(self.http_logger.log_file, 'w').close()
            self._refresh_logs()
            messagebox.showinfo("Success", "Logs cleared successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear logs: {str(e)}")
    
    def _export_logs(self):
        """Export logs to a file"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".log",
                filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if filename:
                if os.path.exists(self.http_logger.log_file):
                    with open(self.http_logger.log_file, 'r', encoding='utf-8') as src:
                        with open(filename, 'w', encoding='utf-8') as dst:
                            dst.write(src.read())
                    messagebox.showinfo("Success", f"Logs exported to: {filename}")
                else:
                    messagebox.showwarning("Warning", "No log file found to export.")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export logs: {str(e)}")
    
    def _open_log_file(self):
        """Open the log file in the default text editor"""
        try:
            log_file_path = self.http_logger.log_file
            
            # Create the log file if it doesn't exist
            if not os.path.exists(log_file_path):
                # Ensure the directory exists
                log_dir = os.path.dirname(log_file_path)
                if not os.path.exists(log_dir):
                    os.makedirs(log_dir)
                
                # Create an empty log file
                with open(log_file_path, 'w', encoding='utf-8') as f:
                    f.write("# Graph API Requests Log\n")
                    f.write(f"# Created: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            
            # Open the file with the default application
            os.startfile(log_file_path)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open log file: {str(e)}")
            print(f"Debug - Log file path: {getattr(self.http_logger, 'log_file', 'Not set')}")
            print(f"Debug - Error: {str(e)}")
    
    def _save_config(self):
        """Save configuration to file"""
        try:
            config = {
                "client_id": self.client_id_var.get(),
                "client_secret": self.client_secret_var.get(),
                "tenant_id": self.tenant_id_var.get(),
                "auth_type": self.auth_type_var.get(),
                "subscription_defaults": {
                    "resource": self.resource_var.get(),
                    "change_type": self.change_type_var.get(),
                    "notification_url": self.notification_url_var.get(),
                    "expiration_hours": self.expiration_var.get(),
                    "include_security_webhooks": self.include_security_var.get()
                }
            }
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Save Configuration"
            )
            
            if filename:
                with open(filename, 'w') as f:
                    json.dump(config, f, indent=2)
                messagebox.showinfo("Success", f"Configuration saved to: {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
    
    def _load_config(self):
        """Load configuration from default file"""
        try:
            config_file = "config.json"
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    config = json.load(f)
                
                # Load authentication config
                self.client_id_var.set(config.get("client_id", ""))
                self.client_secret_var.set(config.get("client_secret", ""))
                self.tenant_id_var.set(config.get("tenant_id", "common"))
                self.auth_type_var.set(config.get("auth_type", "interactive"))
                
                # Load subscription defaults
                self._load_subscription_defaults(config.get("subscription_defaults", {}))
                
        except Exception as e:
            print(f"Could not load default config: {e}")
    
    def _load_subscription_defaults(self, defaults: dict):
        """Load subscription default values into the GUI"""
        try:
            self.resource_var.set(defaults.get("resource", "/me/drive/root"))
            self.change_type_var.set(defaults.get("change_type", "updated"))
            self.notification_url_var.set(defaults.get("notification_url", "https://webhook.site/unique-id"))
            self.expiration_var.set(str(defaults.get("expiration_hours", "24")))
            self.include_security_var.set(defaults.get("include_security_webhooks", True))
            
            print(f"INFO: Loaded subscription defaults from config")
            
        except Exception as e:
            print(f"Could not load subscription defaults: {e}")
    
    def _load_defaults(self):
        """Load subscription defaults from config file"""
        try:
            config_file = "config.json"
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    config = json.load(f)
                
                defaults = config.get("subscription_defaults", {})
                self._load_subscription_defaults(defaults)
                messagebox.showinfo("Success", "Subscription defaults loaded successfully!")
            else:
                messagebox.showwarning("Warning", "No config.json file found")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load defaults: {str(e)}")
    
    def _save_as_defaults(self):
        """Save current subscription field values as defaults to config.json"""
        try:
            config_file = "config.json"
            config = {}
            
            # Load existing config if it exists
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    config = json.load(f)
            
            # Update subscription defaults with current values
            config["subscription_defaults"] = {
                "resource": self.resource_var.get(),
                "change_type": self.change_type_var.get(),
                "notification_url": self.notification_url_var.get(),
                "expiration_hours": self.expiration_var.get(),
                "include_security_webhooks": self.include_security_var.get()
            }
            
            # Save updated config
            with open(config_file, 'w') as f:
                json.dump(config, f, indent=2)
            
            messagebox.showinfo("Success", "Current values saved as defaults!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save as defaults: {str(e)}")
    
    def _reset_subscription_fields(self):
        """Reset subscription fields to application defaults"""
        try:
            self.resource_var.set("/me/drive/root")
            self.change_type_var.set("updated")
            self.notification_url_var.set("https://webhook.site/unique-id")
            self.expiration_var.set("24")
            self.include_security_var.set(True)
            
            messagebox.showinfo("Reset", "Subscription fields reset to defaults!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to reset fields: {str(e)}")
    
    def _show_resource_help(self):
        """Show help dialog for resource patterns"""
        help_text = """Microsoft Graph Resource Patterns

Common resource patterns for subscriptions:

Personal OneDrive:
   /me/drive/root
   /me/drive/items/{item-id}

SharePoint Site:
   /sites/{site-id}/drive/root
   /sites/{site-id}/drive/items/{item-id}

Microsoft 365 Group:
   /groups/{group-id}/drive/root
   /groups/{group-id}/drive/items/{item-id}

Specific User:
   /users/{user-id}/drive/root
   /users/{user-id}/drive/items/{item-id}

Tips:
• Use /me/drive/root for current user's OneDrive
• Replace {site-id} with actual SharePoint site ID
• Replace {group-id} with Microsoft 365 Group ID
• Replace {item-id} with specific file/folder ID
• For security webhooks, monitor root or shared folders

To find IDs:
• Site ID: Graph Explorer → /sites/{hostname}:/sites/{sitename}
• Group ID: Azure AD or Graph Explorer → /groups
• Item ID: Graph Explorer → /me/drive/root/children"""
        
        messagebox.showinfo("Resource Help", help_text)
    
    def _show_webhook_help(self):
        """Show help dialog for webhook URLs"""
        help_text = """Webhook Endpoint Options

Local Testing (with this app):
   http://localhost:8000
   → Use webhook_receiver.py for local testing

Public Testing Services:
   https://webhook.site/unique-id
   → Get a unique URL at webhook.site
   → View notifications in real-time

Azure Functions:
   https://your-app.azurewebsites.net/webhook
   → Deploy a simple webhook receiver

Custom Endpoint:
   https://your-domain.com/webhook
   → Your own webhook implementation

Requirements:
• Must be publicly accessible (not localhost for production)
• Must return HTTP 200 for notifications
• Must respond to validation requests with the token
• HTTPS recommended for production
• No authentication required (secured by subscription)

For Local Testing:
1. Start webhook_receiver.py first
2. Use http://localhost:8000 as URL
3. Or use ngrok to tunnel localhost to public URL

Validation Process:
GET /webhook?validationToken=abc123
Response: abc123 (plain text, HTTP 200)"""
        
        messagebox.showinfo("Webhook Help", help_text)
    
    def _load_config_dialog(self):
        """Load configuration from file dialog"""
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Load Configuration"
            )
            
            if filename:
                with open(filename, 'r') as f:
                    config = json.load(f)
                
                # Load authentication config
                self.client_id_var.set(config.get("client_id", ""))
                self.client_secret_var.set(config.get("client_secret", ""))
                self.tenant_id_var.set(config.get("tenant_id", "common"))
                self.auth_type_var.set(config.get("auth_type", "interactive"))
                
                # Load subscription defaults
                self._load_subscription_defaults(config.get("subscription_defaults", {}))
                
                messagebox.showinfo("Success", f"Configuration loaded from: {filename}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")
    
    def _update_auth_status(self, message: str):
        """Update authentication status display"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.auth_status_text.config(state="normal")
        self.auth_status_text.insert(tk.END, formatted_message)
        self.auth_status_text.config(state="disabled")
        self.auth_status_text.see(tk.END)
    
    def _update_response(self, message: str):
        """Update response display"""
        self.response_text.config(state="normal")
        self.response_text.delete(1.0, tk.END)
        self.response_text.insert(tk.END, message)
        self.response_text.config(state="disabled")
    
    def _update_subscriptions(self, message: str):
        """Update subscriptions display"""
        self.subscriptions_text.config(state="normal")
        self.subscriptions_text.delete(1.0, tk.END)
        self.subscriptions_text.insert(tk.END, message)
        self.subscriptions_text.config(state="disabled")
    
    def _update_webhook_file_list(self, combo_widget):
        """Update webhook file list in combobox"""
        try:
            # Find all webhook notification files
            webhook_files = []
            # Use absolute path to ensure we always read from the correct location
            script_dir = os.path.dirname(os.path.abspath(__file__))
            notifications_dir = os.path.join(script_dir, "webhook_notifications")
            
            if os.path.exists(notifications_dir):
                for filename in os.listdir(notifications_dir):
                    if filename.startswith("webhook_notification_") and filename.endswith(".json"):
                        webhook_files.append(filename)
            
            # Sort by timestamp (newest first)
            webhook_files.sort(reverse=True)
            
            combo_widget['values'] = webhook_files
            if webhook_files:
                combo_widget.set(webhook_files[0])  # Select newest file
                
        except Exception as e:
            print(f"Error updating webhook file list: {e}")
    
    def _browse_webhook_file(self):
        """Browse for webhook notification file"""
        try:
            file_path = filedialog.askopenfilename(
                title="Select Webhook Notification File",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                initialdir=os.getcwd()
            )
            
            if file_path:
                # Get just the filename if in current directory
                if os.path.dirname(file_path) == os.getcwd():
                    file_path = os.path.basename(file_path)
                
                self.webhook_file_var.set(file_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to browse for file: {str(e)}")
    
    def _analyze_latest_webhook(self):
        """Analyze the latest webhook notification (excluding test notifications)"""
        try:
            # Find the latest webhook file
            # Use absolute path to ensure we always read from the correct location
            script_dir = os.path.dirname(os.path.abspath(__file__))
            notifications_dir = os.path.join(script_dir, "webhook_notifications")
            webhook_files = []
            
            if os.path.exists(notifications_dir):
                for filename in os.listdir(notifications_dir):
                    if filename.startswith("webhook_notification_") and filename.endswith(".json"):
                        # Read the file to check if it's a real webhook (not test)
                        file_path = os.path.join(notifications_dir, filename)
                        try:
                            with open(file_path, 'r', encoding='utf-8') as f:
                                data = json.load(f)
                                notification = data.get('notification', {}).get('value', [{}])[0]
                                resource = notification.get('resource', '')
                                # Only include real webhooks (not test ones)
                                if resource != 'test' and 'drives/' in resource:
                                    webhook_files.append((filename, os.path.getmtime(file_path)))
                        except Exception as e:
                            self.logger.debug(f"Could not read webhook file {filename}: {e}")
                            continue
            
            if not webhook_files:
                self._update_change_details("No real webhook notification files found (only test notifications available).")
                return
            
            # Sort by modification time (most recent first) instead of filename
            webhook_files.sort(key=lambda x: x[1], reverse=True)
            latest_file = webhook_files[0][0]
            
            self._update_change_details(f"Analyzing latest real webhook file: {latest_file}")
            
            # Analyze the file
            self._analyze_webhook_file(latest_file)
            
        except Exception as e:
            self._update_change_details(f"Error analyzing latest webhook: {str(e)}")
    
    def _analyze_selected_webhook(self):
        """Analyze the selected webhook notification file"""
        try:
            selected_file = self.webhook_file_var.get().strip()
            if not selected_file:
                messagebox.showwarning("Warning", "Please select a webhook file to analyze.")
                return
            
            self._update_change_details(f"Analyzing selected webhook file: {selected_file}")
            self._analyze_webhook_file(selected_file)
            
        except Exception as e:
            self._update_change_details(f"Error analyzing selected webhook: {str(e)}")
    
    def _analyze_webhook_file(self, filename: str):
        """Analyze a specific webhook notification file using enhanced tracker"""
        try:
            if not EnhancedChangeTracker:
                self._update_change_details("ERROR: Enhanced change tracker not available. Please ensure enhanced_change_tracker.py is in the same directory.")
                return
            
            # Build full file path using absolute path
            script_dir = os.path.dirname(os.path.abspath(__file__))
            notifications_dir = os.path.join(script_dir, "webhook_notifications")
            file_path = os.path.join(notifications_dir, filename)
            
            if not os.path.exists(file_path):
                self._update_change_details(f"ERROR: File not found: {filename}")
                return
            
            self._update_change_details(f"INFO: Processing webhook file: {filename}\n")
            
            # Use the enhanced tracker to analyze changes
            def analyze_worker():
                try:
                    # Load webhook notification
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    notification_data = data.get("notification", {})
                    if "value" not in notification_data:
                        self.root.after(0, lambda: self._update_change_details(f"ERROR: No notifications found in {filename}"))
                        return
                    
                    # Initialize enhanced tracker with HTTP logger
                    tracker = EnhancedChangeTracker(http_logger=self.http_logger)
                    
                    # Set the webhook timestamp for accurate correlation
                    webhook_timestamp = data.get("timestamp")
                    if webhook_timestamp:
                        tracker._current_webhook_timestamp = webhook_timestamp
                    
                    notifications = notification_data["value"]
                    
                    all_analyses = []
                    for i, notification in enumerate(notifications, 1):
                        print(f"Analyzing notification {i}...")
                        analysis = tracker.process_webhook_notification(notification)
                        all_analyses.append(analysis)
                    
                    # Format the results for display
                    if all_analyses:
                        details_text = f"Analysis complete for {filename}\n\n"
                        details_text += f"Analyzed {len(all_analyses)} notification(s):\n\n"
                        
                        for i, analysis in enumerate(all_analyses, 1):
                            if "error" in analysis:
                                details_text += f"--- Notification {i} ---\n"
                                details_text += f"ERROR: {analysis['error']}\n\n"
                                continue
                            
                            # Extract key information
                            webhook_info = analysis.get('webhook_notification', {})
                            item_details = analysis.get('item_details', {})
                            activities = item_details.get('activities', [])
                            permissions = analysis.get('permissions', [])
                            analysis_summary = analysis.get('analysis_summary', [])
                            webhook_correlation = analysis.get('webhook_correlation', {})
                            
                            details_text += f"--- Notification {i} ---\n"
                            details_text += f"Resource: {webhook_info.get('resource', 'N/A')}\n"
                            details_text += f"Change Type: {webhook_info.get('changeType', 'N/A')}\n"
                            
                            # Show the new analysis summary first (most important!)
                            if analysis_summary:
                                details_text += f"\nENHANCED ANALYSIS:\n"
                                for summary_item in analysis_summary:
                                    category = summary_item.get('category', 'Analysis')
                                    details = summary_item.get('details', 'No details')
                                    details_text += f"  • {category}: {details}\n"
                            
                            # Show webhook correlation details if available
                            if webhook_correlation:
                                confidence = webhook_correlation.get('correlation_confidence', 'unknown')
                                details_text += f"\nCORRELATION ANALYSIS:\n"
                                details_text += f"  Matched Item: {webhook_correlation.get('matched_item', 'Unknown')}\n"
                                details_text += f"  Change Type: {webhook_correlation.get('change_type', 'Unknown')}\n"
                                
                                # Show operation details if available
                                operation_details = webhook_correlation.get('operation_details', '')
                                if operation_details:
                                    details_text += f"  Operation: {operation_details}\n"
                                
                                details_text += f"  Change Time: {webhook_correlation.get('change_time', 'Unknown')}\n"
                                details_text += f"  Webhook Time: {webhook_correlation.get('webhook_time', 'Unknown')}\n"
                                details_text += f"  Latency: {webhook_correlation.get('latency_seconds', 0):.1f} seconds\n"
                                details_text += f"  Confidence: {confidence}\n"
                            
                            if item_details:
                                details_text += f"Item: {item_details.get('name', 'Unknown')}\n"
                                details_text += f"Last Modified: {item_details.get('lastModifiedDateTime', 'N/A')}\n"
                                details_text += f"Size: {item_details.get('size', 'N/A')} bytes\n"
                            
                            # Show recent activities (most important!)
                            if activities:
                                details_text += f"\n🎯 RECENT ACTIVITIES ({len(activities)} total):\n"
                                
                                security_activities = []
                                for j, activity in enumerate(activities[:10], 1):  # Show top 10
                                    action = activity.get('action', {})
                                    actor = activity.get('actor', {}).get('user', {})
                                    time_info = activity.get('times', {})
                                    
                                    # Determine activity type
                                    if 'share' in action:
                                        details_text += f"  {j}. 🔒 SHARE (Security-Related)\n"
                                        recipients = action['share'].get('recipients', [])
                                        recipient_names = []
                                        for recipient in recipients:
                                            user_info = recipient.get('user', {})
                                            name = user_info.get('displayName', user_info.get('email', 'Unknown'))
                                            recipient_names.append(name)
                                        details_text += f"     👥 Shared with: {', '.join(recipient_names)}\n"
                                        security_activities.append(f"Share with {', '.join(recipient_names)}")
                                    
                                    elif 'rename' in action:
                                        old_name = action['rename'].get('oldName', 'Unknown')
                                        details_text += f"  {j}. 📝 RENAME\n"
                                        details_text += f"     📄 From: {old_name}\n"
                                    
                                    elif 'create' in action:
                                        details_text += f"  {j}. ➕ CREATE\n"
                                        details_text += f"     📄 New item created\n"
                                    
                                    elif 'edit' in action:
                                        details_text += f"  {j}. ✏️ EDIT\n"
                                        if 'version' in action:
                                            version = action['version'].get('newVersion', 'Unknown')
                                            details_text += f"     📋 New version: {version}\n"
                                    
                                    # Add actor and time info
                                    details_text += f"     👤 By: {actor.get('displayName', 'Unknown')}\n"
                                    recorded_time = time_info.get('recordedDateTime', 'Unknown')
                                    try:
                                        if recorded_time != 'Unknown':
                                            dt = datetime.fromisoformat(recorded_time.replace('Z', '+00:00'))
                                            formatted_time = dt.strftime('%Y-%m-%d %H:%M:%S UTC')
                                        else:
                                            formatted_time = 'Unknown'
                                    except:
                                        formatted_time = recorded_time
                                    details_text += f"     ⏰ When: {formatted_time}\n"
                                
                                # Highlight security activities
                                if security_activities:
                                    details_text += f"\n🔒 SECURITY SUMMARY:\n"
                                    details_text += f"   Found {len(security_activities)} security-related activity/activities!\n"
                                    for sec_activity in security_activities:
                                        details_text += f"   • {sec_activity}\n"
                            
                            # Show permissions
                            if permissions:
                                details_text += f"\n� CURRENT PERMISSIONS ({len(permissions)}):\n"
                                for j, perm in enumerate(permissions[:5], 1):  # Show top 5
                                    if 'link' in perm:
                                        link_info = perm['link']
                                        link_type = link_info.get('type', 'unknown')
                                        scope = link_info.get('scope', 'unknown')
                                        details_text += f"  {j}. 🔗 Link: {link_type} ({scope})\n"
                                    elif 'grantedTo' in perm:
                                        granted_to = perm['grantedTo']
                                        if 'user' in granted_to:
                                            user_name = granted_to['user'].get('displayName', 'Unknown')
                                            details_text += f"  {j}. 👤 User: {user_name}\n"
                                        elif 'group' in granted_to:
                                            group_name = granted_to['group'].get('displayName', 'Unknown')
                                            details_text += f"  {j}. 👥 Group: {group_name}\n"
                            
                            details_text += "\n"
                        
                        self.root.after(0, lambda: self._update_change_details(details_text))
                        
                        # Play sound for successful analysis
                        if self.sound_manager:
                            self.sound_manager.play_success_sound()
                    else:
                        self.root.after(0, lambda: self._update_change_details(f"❌ No analysis results for {filename}"))
                
                except Exception as e:
                    error_msg = f"❌ Error during analysis: {str(e)}\n\n"
                    error_msg += "💡 This might happen if:\n"
                    error_msg += "• Authentication has expired\n"
                    error_msg += "• Network connectivity issues\n"
                    error_msg += "• The resource doesn't support enhanced tracking\n"
                    self.root.after(0, lambda: self._update_change_details(error_msg))
            
            # Run analysis in background thread
            threading.Thread(target=analyze_worker, daemon=True).start()
            
        except Exception as e:
            self._update_change_details(f"❌ Error setting up analysis: {str(e)}")
    
    def _refresh_changes(self):
        """Refresh the change details display"""
        try:
            # Look for existing change detail files
            change_analysis_dir = "change_analysis"
            change_files = []
            
            if os.path.exists(change_analysis_dir):
                for filename in os.listdir(change_analysis_dir):
                    if (filename.startswith("enhanced_analysis_") or filename.startswith("change_details_")) and filename.endswith(".json"):
                        change_files.append(filename)
            
            if not change_files:
                self._update_change_details("No change analysis files found yet.\n\n"
                                          "💡 To generate detailed change analysis:\n"
                                          "1. Ensure you have webhook notifications in the webhook_notifications folder\n"
                                          "2. Click 'Analyze Latest Webhook' or select a specific file\n"
                                          "3. The system will use Microsoft Graph Enhanced Tracker to get detailed changes")
                return
            
            # Sort by timestamp (newest first)
            change_files.sort(reverse=True)
            
            # Display summary of available change files
            summary = f"📊 Found {len(change_files)} change analysis file(s):\n\n"
            
            for i, filename in enumerate(change_files[:5], 1):  # Show latest 5
                try:
                    file_path = os.path.join(change_analysis_dir, filename)
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    timestamp = data.get('timestamp', 'Unknown')
                    
                    # Check for different file formats
                    if 'analysis_summary' in data:
                        summary_count = len(data.get('analysis_summary', []))
                        summary += f"{i}. {filename}\n"
                        summary += f"   📅 {timestamp}\n"
                        summary += f"   📊 {summary_count} insights\n\n"
                    elif 'total_changes' in data:
                        change_count = data.get('total_changes', 0)
                        summary += f"{i}. {filename}\n"
                        summary += f"   📅 {timestamp}\n"
                        summary += f"   📊 {change_count} detailed changes\n\n"
                    else:
                        summary += f"{i}. {filename}\n"
                        summary += f"   📅 {timestamp}\n"
                        summary += f"   📊 Analysis file\n\n"
                    
                except Exception as e:
                    summary += f"{i}. {filename} (Error reading: {e})\n\n"
            
            if len(change_files) > 5:
                summary += f"... and {len(change_files) - 5} more files\n\n"
            
            summary += "💡 Use 'Analyze Latest Webhook' to create new detailed change analysis."
            
            self._update_change_details(summary)
            
        except Exception as e:
            self._update_change_details(f"Error refreshing changes: {str(e)}")
    
    def _clear_change_analysis(self):
        """Clear the change analysis display"""
        self._update_change_details("Change analysis cleared.\n\n"
                                  "Use 'Analyze Latest Webhook' or select a specific webhook file to analyze changes.")
    
    def _update_change_details(self, message: str):
        """Update change details display"""
        self.change_details_text.config(state="normal")
        self.change_details_text.delete(1.0, tk.END)
        self.change_details_text.insert(tk.END, message)
        self.change_details_text.config(state="disabled")
        self.change_details_text.see(tk.END)
    
    def _refresh_subscription_dropdown(self):
        """Refresh the subscription dropdown list"""
        try:
            if not self.subscription_manager:
                messagebox.showwarning("Warning", "Please authenticate first.")
                return
            
            def refresh_worker():
                try:
                    result = self.subscription_manager.list_subscriptions()
                    
                    if result.get('success'):
                        subscription_data = result.get('data', {})
                        subscriptions = subscription_data.get('value', [])
                        
                        subscription_options = []
                        for sub in subscriptions:
                            sub_id = sub.get('id', 'Unknown')
                            resource = sub.get('resource', 'Unknown')
                            # Truncate long resources for display
                            display_resource = resource if len(resource) <= 50 else resource[:47] + "..."
                            subscription_options.append(f"{sub_id} - {display_resource}")
                        
                        # Update dropdown in main thread
                        self.root.after(0, lambda: self._update_subscription_dropdown(subscription_options))
                    else:
                        error_msg = result.get('message', 'Unknown error occurred')
                        self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to get subscriptions: {error_msg}"))
                        
                except Exception as e:
                    error_msg = str(e)
                    self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to refresh subscriptions: {error_msg}"))
            
            # Run in background thread
            threading.Thread(target=refresh_worker, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh subscription dropdown: {str(e)}")
    
    def _update_subscription_dropdown(self, options):
        """Update subscription dropdown with options"""
        self.subscription_combo['values'] = options
        if options:
            self.subscription_combo.set(options[0])
        else:
            self.subscription_combo.set("")
    
    def _delete_selected_subscription(self):
        """Delete the selected subscription"""
        try:
            if not self.subscription_manager:
                messagebox.showwarning("Warning", "Please authenticate first.")
                return
            
            selected = self.subscription_id_var.get()
            if not selected:
                messagebox.showwarning("Warning", "Please select a subscription to delete.")
                return
            
            # Extract subscription ID from the display text
            subscription_id = selected.split(' - ')[0]
            
            # Confirm deletion
            result = messagebox.askyesno(
                "Confirm Deletion",
                f"Are you sure you want to delete subscription:\n{subscription_id}?\n\nThis action cannot be undone."
            )
            
            if not result:
                return
            
            def delete_worker():
                try:
                    success = self.subscription_manager.delete_subscription(subscription_id)
                    
                    if success:
                        self.root.after(0, lambda: messagebox.showinfo("Success", f"Subscription {subscription_id} deleted successfully."))
                        self.root.after(0, self._refresh_subscriptions)
                        self.root.after(0, self._refresh_subscription_dropdown)
                        
                        if self.sound_manager:
                            self.sound_manager.play_success_sound()
                    else:
                        self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to delete subscription {subscription_id}"))
                        
                except Exception as e:
                    self.root.after(0, lambda: messagebox.showerror("Error", f"Error deleting subscription: {str(e)}"))
            
            # Run in background thread
            threading.Thread(target=delete_worker, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete subscription: {str(e)}")
    
    def _delete_all_subscriptions(self):
        """Delete all subscriptions"""
        try:
            if not self.subscription_manager:
                messagebox.showwarning("Warning", "Please authenticate first.")
                return
            
            # Get current subscriptions count
            def count_worker():
                try:
                    subscriptions = self.subscription_manager.list_subscriptions()
                    count = len(subscriptions) if subscriptions else 0
                    
                    if count == 0:
                        self.root.after(0, lambda: messagebox.showinfo("Info", "No subscriptions found to delete."))
                        return
                    
                    # Confirm deletion in main thread
                    self.root.after(0, lambda: self._confirm_delete_all(count))
                    
                except Exception as e:
                    self.root.after(0, lambda: messagebox.showerror("Error", f"Error getting subscriptions: {str(e)}"))
            
            # Run in background thread
            threading.Thread(target=count_worker, daemon=True).start()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete all subscriptions: {str(e)}")
    
    def _confirm_delete_all(self, count):
        """Confirm deletion of all subscriptions"""
        result = messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to delete ALL {count} subscription(s)?\n\nThis action cannot be undone."
        )
        
        if not result:
            return
        
        def delete_all_worker():
            try:
                subscriptions = self.subscription_manager.list_subscriptions()
                
                if not subscriptions:
                    self.root.after(0, lambda: messagebox.showinfo("Info", "No subscriptions found to delete."))
                    return
                
                deleted_count = 0
                failed_count = 0
                
                for subscription in subscriptions:
                    sub_id = subscription.get('id', '')
                    if sub_id:
                        try:
                            success = self.subscription_manager.delete_subscription(sub_id)
                            if success:
                                deleted_count += 1
                            else:
                                failed_count += 1
                        except:
                            failed_count += 1
                
                # Show results
                if failed_count == 0:
                    self.root.after(0, lambda: messagebox.showinfo("Success", f"Successfully deleted all {deleted_count} subscription(s)."))
                    if self.sound_manager:
                        self.sound_manager.play_success_sound()
                else:
                    self.root.after(0, lambda: messagebox.showwarning("Partial Success", f"Deleted {deleted_count} subscription(s), failed to delete {failed_count}."))
                
                # Refresh displays
                self.root.after(0, self._refresh_subscriptions)
                self.root.after(0, self._refresh_subscription_dropdown)
                
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Error deleting subscriptions: {str(e)}"))
        
        # Run in background thread
        threading.Thread(target=delete_all_worker, daemon=True).start()
    
    def run(self):
        """Run the GUI application"""
        self.root.mainloop()

def main():
    """Main function"""
    print("Starting Microsoft Graph Security Webhook Tester...")
    
    try:
        app = GraphWebhookTesterGUI()
        app.run()
    except Exception as e:
        print(f"Failed to start application: {e}")
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()