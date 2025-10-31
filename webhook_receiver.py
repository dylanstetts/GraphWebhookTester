#!/usr/bin/env python3
"""
Simple webhook receiver for testing Microsoft Graph notifications.
This creates a local HTTP server to receive webhook notifications.
"""

import http.server
import socketserver
import json
import urllib.parse
import os
from datetime import datetime
import threading
import webbrowser

class WebhookHandler(http.server.BaseHTTPRequestHandler):
    """HTTP request handler for webhook notifications"""
    
    def do_GET(self):
        """Handle GET requests (validation)"""
        # Parse query parameters
        parsed_url = urllib.parse.urlparse(self.path)
        query_params = urllib.parse.parse_qs(parsed_url.query)
        
        # Log the request
        self.log_request_details("GET")
        
        # Check for validation token
        if 'validationToken' in query_params:
            validation_token = query_params['validationToken'][0]
            
            print(f"INFO: Webhook validation request received!")
            print(f"INFO: Validation token: {validation_token}")
            
            # Respond with validation token in plain text
            self.send_response(200)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write(validation_token.encode('utf-8'))
            
            print("INFO: Validation response sent successfully!")
        else:
            # Regular GET request
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            
            html = """
            <html>
            <head><title>Webhook Receiver</title></head>
            <body>
            <h1>Microsoft Graph Webhook Receiver</h1>
            <p>This server is ready to receive webhook notifications.</p>
            <p>Current time: {}</p>
            </body>
            </html>
            """.format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            
            self.wfile.write(html.encode('utf-8'))
    
    def do_POST(self):
        """Handle POST requests (notifications and validation)"""
        self.log_request_details("POST")
        
        # Parse query parameters for validation token
        parsed_url = urllib.parse.urlparse(self.path)
        query_params = urllib.parse.parse_qs(parsed_url.query)
        
        # Check for validation token in POST request (some webhooks send POST for validation)
        if 'validationToken' in query_params:
            validation_token = query_params['validationToken'][0]
            
            print(f"Webhook validation request received via POST!")
            print(f"Validation token: {validation_token}")
            
            # Respond with validation token in plain text
            self.send_response(200)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write(validation_token.encode('utf-8'))
            
            print("Validation response sent successfully!")
            return
        
        # Read the request body for normal notifications
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length)
        
        # Handle empty body (validation requests sometimes have empty body)
        if content_length == 0 or not post_data.strip():
            print("Empty POST request received (likely validation)")
            self.send_response(200)
            self.send_header('Content-type', 'text/plain')
            self.end_headers()
            self.wfile.write(b"OK")
            return
        
        try:
            # Parse JSON payload
            notification_data = json.loads(post_data.decode('utf-8'))
            
            print(f"\nWEBHOOK NOTIFICATION RECEIVED!")
            print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Raw payload:")
            print(json.dumps(notification_data, indent=2))
            
            # Process notifications
            if 'value' in notification_data:
                notifications = notification_data['value']
                print(f"\nFound {len(notifications)} notification(s):")
                
                for i, notification in enumerate(notifications, 1):
                    print(f"\n--- Notification {i} ---")
                    print(f"Subscription ID: {notification.get('subscriptionId', 'N/A')}")
                    print(f"Change Type: {notification.get('changeType', 'N/A')}")
                    print(f"Resource: {notification.get('resource', 'N/A')}")
                    print(f"Client State: {notification.get('clientState', 'N/A')}")
                    print(f"Subscription Expiration: {notification.get('subscriptionExpirationDateTime', 'N/A')}")
                    
                    # Check for resource data
                    if 'resourceData' in notification:
                        resource_data = notification['resourceData']
                        print(f"Resource Data:")
                        print(f"   ID: {resource_data.get('id', 'N/A')}")
                        print(f"   Type: {resource_data.get('@odata.type', 'N/A')}")
                        print(f"   ETag: {resource_data.get('@odata.etag', 'N/A')}")
            
            # Log to file
            self.log_to_file(notification_data)
            
            # Send success response
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            
            response = {"status": "received", "timestamp": datetime.now().isoformat()}
            self.wfile.write(json.dumps(response).encode('utf-8'))
            
            print("Notification processed successfully!")
            
        except json.JSONDecodeError as e:
            print(f"Error parsing JSON: {e}")
            print(f"Raw data: {post_data}")
            
            self.send_response(400)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            
            error_response = {"error": "Invalid JSON", "details": str(e)}
            self.wfile.write(json.dumps(error_response).encode('utf-8'))
        
        except Exception as e:
            print(f"Error processing notification: {e}")
            
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            
            error_response = {"error": "Processing failed", "details": str(e)}
            self.wfile.write(json.dumps(error_response).encode('utf-8'))
    
    def log_request_details(self, method: str):
        """Log request details"""
        print(f"\n{method} request received:")
        print(f"Path: {self.path}")
        print(f"Client: {self.client_address}")
        print(f"Headers:")
        for header, value in self.headers.items():
            print(f"   {header}: {value}")
    
    def log_to_file(self, data: dict):
        """Log notification to file"""
        try:
            # Create webhook_notifications directory if it doesn't exist
            # Use absolute path to ensure we always write to the correct location
            script_dir = os.path.dirname(os.path.abspath(__file__))
            notifications_dir = os.path.join(script_dir, "webhook_notifications")
            if not os.path.exists(notifications_dir):
                os.makedirs(notifications_dir)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"webhook_notification_{timestamp}.json"
            filepath = os.path.join(notifications_dir, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump({
                    "timestamp": datetime.now().isoformat(),
                    "notification": data
                }, f, indent=2)
            
            print(f"Notification saved to: {filepath}")
            
        except Exception as e:
            print(f"Error saving notification to file: {e}")
    
    def log_message(self, format, *args):
        """Override to reduce server logging noise"""
        pass

def start_webhook_server(port: int = 8000):
    """Start the webhook server"""
    
    print("Starting Microsoft Graph Webhook Receiver...")
    print(f"Server will run on: http://localhost:{port}")
    print(f"Webhook URL: http://localhost:{port}")
    print("Press Ctrl+C to stop the server\n")
    
    try:
        with socketserver.TCPServer(("", port), WebhookHandler) as httpd:
            print(f"Server started successfully on port {port}")
            
            # Open browser to show the server is running
            def open_browser():
                import time
                time.sleep(1)  # Wait a moment for server to start
                try:
                    webbrowser.open(f"http://localhost:{port}")
                except:
                    pass
            
            threading.Thread(target=open_browser, daemon=True).start()
            
            print("Use this URL as your notification URL in the Graph webhook tester:")
            print(f"   http://localhost:{port}")
            print("\nWaiting for webhook notifications...\n")
            
            httpd.serve_forever()
            
    except KeyboardInterrupt:
        print("\nServer stopped by user")
    except OSError as e:
        if e.errno == 10048:  # Port already in use
            print(f"Port {port} is already in use. Try a different port:")
            print(f"   python webhook_receiver.py --port {port + 1}")
        else:
            print(f"Error starting server: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

def main():
    """Main function"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Microsoft Graph Webhook Receiver")
    parser.add_argument("--port", type=int, default=8000, help="Port to run the server on (default: 8000)")
    
    args = parser.parse_args()
    
    start_webhook_server(args.port)

if __name__ == "__main__":
    main()
