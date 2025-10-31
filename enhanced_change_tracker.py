#!/usr/bin/env python3
"""
Enhanced Change Tracker for Microsoft Graph
Uses direct API calls to get detailed information about what actually changed.
This approach works much better than delta queries for webhook notifications.
"""

import json
import requests
import os
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
import logging

# Import MSAL for authentication
try:
    import msal
except ImportError:
    print("MSAL not installed. Please install it with: pip install msal")
    exit(1)

class EnhancedChangeTracker:
    """Tracks changes using direct Microsoft Graph API calls"""
    
    def __init__(self, config_file: str = "config.json", http_logger=None):
        """Initialize the enhanced change tracker"""
        self.config_file = config_file
        self.access_token = None
        self.http_logger = http_logger
        self.logger = self._setup_logger()
        
        # Load configuration
        self.config = self._load_config()
        
    def _setup_logger(self) -> logging.Logger:
        """Setup logging for change tracker"""
        logger = logging.getLogger("EnhancedChangeTracker")
        logger.setLevel(logging.INFO)
        
        # Ensure logs directory exists
        import os
        os.makedirs("logs", exist_ok=True)
        
        # Create file handler
        handler = logging.FileHandler("logs/enhanced_changes.log", encoding='utf-8')
        handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)
        
        # Add handler to logger
        if not logger.handlers:
            logger.addHandler(handler)
        
        return logger
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                return {}
        except Exception as e:
            self.logger.error(f"Error loading config: {e}")
            return {}
    
    def authenticate(self) -> bool:
        """Authenticate with Microsoft Graph"""
        try:
            if not self.config:
                self.logger.error("No configuration found")
                return False
            
            # Create MSAL app
            app = msal.ConfidentialClientApplication(
                client_id=self.config.get("client_id"),
                client_credential=self.config.get("client_secret"),
                authority=f"https://login.microsoftonline.com/{self.config.get('tenant_id', 'common')}"
            )
            
            # Get access token using client credentials flow
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.logger.info("Successfully authenticated with Microsoft Graph")
                return True
            else:
                self.logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
                
        except Exception as e:
            self.logger.error(f"Authentication error: {e}")
            return False
    
    def _make_graph_request(self, url: str, method: str = "GET", data: Dict = None) -> Optional[Dict]:
        """Make authenticated request to Microsoft Graph"""
        if not self.access_token:
            if not self.authenticate():
                return None

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        # Log the request if HTTP logger is available
        if self.http_logger:
            if method == "GET":
                self.http_logger.log_request(method, url, headers)
            elif method == "POST":
                body = json.dumps(data) if data else None
                self.http_logger.log_request(method, url, headers, body)

        try:
            if method == "GET":
                response = requests.get(url, headers=headers)
            elif method == "POST":
                response = requests.post(url, headers=headers, json=data)
            else:
                raise ValueError(f"Unsupported method: {method}")

            # Log the response if HTTP logger is available
            if self.http_logger:
                response_body = response.text
                self.http_logger.log_response(response.status_code, response.headers, response_body)

            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 401:
                # Token might be expired, try to refresh
                if self.authenticate():
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    
                    # Log retry request
                    if self.http_logger:
                        if method == "GET":
                            self.http_logger.log_request(method, url, headers)
                        elif method == "POST":
                            body = json.dumps(data) if data else None
                            self.http_logger.log_request(method, url, headers, body)
                    
                    if method == "GET":
                        response = requests.get(url, headers=headers)
                    elif method == "POST":
                        response = requests.post(url, headers=headers, json=data)
                    
                    # Log retry response
                    if self.http_logger:
                        response_body = response.text
                        self.http_logger.log_response(response.status_code, response.headers, response_body)
                    
                    response.raise_for_status()
                    return response.json()
            
            self.logger.error(f"HTTP error: {e}")
            return None
        except Exception as e:
            self.logger.error(f"Request error: {e}")
            return None

    def get_item_details(self, resource_path: str) -> Optional[Dict[str, Any]]:
        """Get detailed information about a specific item (comprehensive approach)"""
        try:
            # Convert resource path to API URL
            if resource_path.startswith("/"):
                resource_path = resource_path[1:]  # Remove leading slash
            
            # Build the API URL - try to get as much information as possible
            api_url = f"https://graph.microsoft.com/v1.0/{resource_path}"
            
            # Use a more comprehensive expand that's more likely to work
            expand_params = [
                "permissions",
                "children($top=10)",  # Get some child items
                "lastModifiedBy"      # Get who last modified it
            ]
            
            api_url += f"?$expand={','.join(expand_params)}"
            
            self.logger.info(f"Getting item details from: {api_url}")
            
            response = self._make_graph_request(api_url)
            
            # If the expand fails, try without expand
            if not response:
                self.logger.info("Expand failed, trying simple request...")
                simple_url = f"https://graph.microsoft.com/v1.0/{resource_path}"
                response = self._make_graph_request(simple_url)
            
            return response
            
        except Exception as e:
            self.logger.error(f"Error getting item details: {e}")
            return None
    
    def get_item_permissions(self, resource_path: str) -> List[Dict[str, Any]]:
        """Get permissions for a specific item"""
        try:
            if resource_path.startswith("/"):
                resource_path = resource_path[1:]
            
            permissions_url = f"https://graph.microsoft.com/v1.0/{resource_path}/permissions"
            
            self.logger.info(f"Getting permissions from: {permissions_url}")
            
            response = self._make_graph_request(permissions_url)
            if response and 'value' in response:
                return response['value']
            
            return []
            
        except Exception as e:
            self.logger.error(f"Error getting permissions: {e}")
            return []
    
    def get_item_activities(self, resource_path: str, limit: int = 10) -> List[Dict[str, Any]]:
        """Get recent activities for a specific item"""
        try:
            if resource_path.startswith("/"):
                resource_path = resource_path[1:]
            
            activities_url = f"https://graph.microsoft.com/v1.0/{resource_path}/activities?$top={limit}&$orderby=lastModifiedDateTime desc"
            
            self.logger.info(f"Getting activities from: {activities_url}")
            
            response = self._make_graph_request(activities_url)
            if response and 'value' in response:
                return response['value']
            
            return []
            
        except Exception as e:
            self.logger.error(f"Error getting activities: {e}")
            return []
    
    def get_item_versions(self, resource_path: str) -> List[Dict[str, Any]]:
        """Get versions for a specific item"""
        try:
            if resource_path.startswith("/"):
                resource_path = resource_path[1:]
            
            versions_url = f"https://graph.microsoft.com/v1.0/{resource_path}/versions"
            
            self.logger.info(f"Getting versions from: {versions_url}")
            
            response = self._make_graph_request(versions_url)
            if response and 'value' in response:
                return response['value']
            
            return []
            
        except Exception as e:
            self.logger.error(f"Error getting versions: {e}")
            return []
    
    def _analyze_specific_changes(self, analysis: Dict, item_details: Optional[Dict], webhook_time: datetime):
        """Analyze what specifically changed based on activities, versions, and recent modifications"""
        
        # Safety check for None item_details
        if item_details is None:
            analysis["analysis_summary"].append({
                "category": "Specific Changes",
                "details": "[ERROR] Cannot analyze changes - no item details available",
                "correlation_method": "failed_api_call"
            })
            return
        
        # Check recent activities for specific change details
        activities = item_details.get('activities', {}).get('value', [])
        if activities:
            recent_changes = []
            for activity in activities[:5]:  # Check top 5 activities
                action = activity.get('action', {})
                actor = activity.get('actor', {}).get('user', {})
                times = activity.get('times', {})
                
                # Analyze the type of change
                action_type = action.get('@odata.type', '')
                if 'Create' in action_type:
                    change_detail = f"Created: {action.get('name', 'Unknown item')}"
                elif 'Delete' in action_type:
                    change_detail = f"Deleted: {action.get('name', 'Unknown item')}"
                elif 'Edit' in action_type:
                    change_detail = f"Modified: {action.get('name', 'Unknown item')}"
                elif 'Rename' in action_type:
                    change_detail = f"Renamed: {action.get('oldName', 'Unknown')} → {action.get('newName', 'Unknown')}"
                elif 'Share' in action_type:
                    change_detail = f"Shared: {action.get('name', 'Unknown item')}"
                else:
                    change_detail = f"Action: {action_type}"
                
                recent_changes.append({
                    "change": change_detail,
                    "user": actor.get('displayName', 'Unknown user'),
                    "timestamp": times.get('recordedDateTime', 'Unknown time'),
                    "action_type": action_type
                })
            
            if recent_changes:
                analysis["specific_changes"] = recent_changes
                analysis["analysis_summary"].append({
                    "category": "Specific Changes Detected",
                    "details": f"[CHANGES] Found {len(recent_changes)} recent changes: {recent_changes[0]['change']}",
                    "is_security_related": any('Share' in change['action_type'] for change in recent_changes),
                    "correlation_strength": "HIGH",
                    "correlation_method": "activity_analysis"
                })
        
        # Check file versions for content changes
        versions = item_details.get('versions', {}).get('value', [])
        if versions and len(versions) > 1:
            latest_version = versions[0]
            previous_version = versions[1] if len(versions) > 1 else None
            
            if latest_version and previous_version:
                latest_modified = latest_version.get('lastModifiedDateTime', '')
                previous_modified = previous_version.get('lastModifiedDateTime', '')
                modified_by = latest_version.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'Unknown')
                
                analysis["analysis_summary"].append({
                    "category": "File Version Changes",
                    "details": f"[VERSION] File updated by {modified_by} - Latest: {latest_modified}",
                    "correlation_strength": "HIGH",
                    "correlation_method": "version_analysis"
                })
        
        # Check for recent folder content changes (if this is a folder)
        if item_details.get('@odata.type') == '#microsoft.graph.driveItem' and not item_details.get('file'):
            # This is likely a folder - check for recent modifications
            last_modified = item_details.get('lastModifiedDateTime')
            if last_modified:
                try:
                    modified_time = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                    time_diff = abs((webhook_time - modified_time.replace(tzinfo=None)).total_seconds())
                    
                    if time_diff <= 300:  # Within 5 minutes (more reasonable for folder changes)
                        analysis["analysis_summary"].append({
                            "category": "Folder Content Change",
                            "details": f"[FOLDER] Folder contents recently modified ({time_diff:.0f}s ago)",
                            "correlation_strength": "HIGH",
                            "correlation_method": "folder_analysis"
                        })
                except Exception as e:
                    self.logger.warning(f"Could not parse folder modification time: {e}")

    def _get_recent_drive_changes(self, drive_id: str) -> List[Dict]:
        """Get recent changes in a drive to understand what triggered the webhook"""
        try:
            # Note: The /recent endpoint may not be available for all drives
            # Try to get recent items, but handle failures gracefully
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/recent"
            recent_items = self._make_graph_request(url)
            
            if recent_items and 'value' in recent_items:
                # Filter to very recent items (last 10 minutes)
                from datetime import datetime, timedelta
                cutoff_time = datetime.now() - timedelta(minutes=10)
                
                recent_changes = []
                for item in recent_items['value'][:20]:  # Check top 20 recent items
                    last_modified = item.get('lastModifiedDateTime')
                    if last_modified:
                        try:
                            modified_time = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                            if modified_time.replace(tzinfo=None) > cutoff_time:
                                recent_changes.append({
                                    "name": item.get('name', 'Unknown'),
                                    "lastModifiedDateTime": last_modified,
                                    "lastModifiedBy": item.get('lastModifiedBy', {}),
                                    "size": item.get('size'),
                                    "webUrl": item.get('webUrl'),
                                    "changeType": "modified"  # Inferred from recent list
                                })
                        except Exception:
                            continue
                
                return recent_changes[:10]  # Return top 10 recent changes
            
            return []
            
        except Exception as e:
            self.logger.warning(f"Could not get recent drive changes: {e}")
            return []

    def get_recent_activities(self, resource: str, minutes: int = 2) -> List[Dict]:
        """Get recent activities for correlation with webhook notifications"""
        try:
            # Try to get drive activities from audit logs
            # This is a simplified approach - in practice you might need different endpoints
            
            # Extract drive ID from resource
            if '/drives/' in resource:
                parts = resource.split('/drives/')
                if len(parts) > 1:
                    drive_part = parts[1].split('/')[0]
                    
                    # Try to get recent items from the drive (fallback approach)
                    url = f"https://graph.microsoft.com/v1.0/drives/{drive_part}/root/children?$top=10"
                    recent_items = self._make_graph_request(url)
                    
                    if recent_items and 'value' in recent_items:
                        return recent_items['value'][:5]  # Return top 5 recent items
            
            return []
            
        except Exception as e:
            self.logger.warning(f"Could not get recent activities: {e}")
            return []

    def _get_item_activities(self, resource: str) -> List[Dict]:
        """Get activities for a specific item using alternative approaches"""
        try:
            # Convert resource path to API URL
            if resource.startswith("/"):
                resource = resource[1:]  # Remove leading slash
            
            # Try different approaches since direct activities endpoint may fail
            
            # Approach 1: Try to get activities directly (may fail with 400)
            try:
                activities_url = f"https://graph.microsoft.com/v1.0/{resource}/activities?$top=5"
                activities_response = self._make_graph_request(activities_url)
                if activities_response and 'value' in activities_response:
                    return activities_response['value']
            except Exception as e:
                self.logger.debug(f"Direct activities call failed: {e}")
            
            # Approach 2: Try to get children (for folders) and check their modifications
            try:
                if '/root' in resource:  # This is a folder/drive root
                    children_url = f"https://graph.microsoft.com/v1.0/{resource}/children?$top=10&$orderby=lastModifiedDateTime desc"
                    children_response = self._make_graph_request(children_url)
                    if children_response and 'value' in children_response:
                        # Convert recent children to activity-like format
                        activities = []
                        for child in children_response['value'][:5]:
                            activities.append({
                                "action": {"type": "recent_modification"},
                                "actor": child.get('lastModifiedBy', {}),
                                "times": {"recordedDateTime": child.get('lastModifiedDateTime')},
                                "target": {"name": child.get('name')}
                            })
                        return activities
            except Exception as e:
                self.logger.debug(f"Children approach failed: {e}")
            
            # If all approaches fail, return empty list
            return []
            
        except Exception as e:
            self.logger.warning(f"Could not get item activities: {e}")
            return []

    def analyze_change_details_with_correlation(self, notification: Dict[str, Any]) -> Dict[str, Any]:
        """Enhanced analysis with multiple correlation strategies"""
        
        # Extract key information from notification
        resource = notification.get('resource', '')
        change_type = notification.get('changeType', '')
        subscription_id = notification.get('subscriptionId', '')
        
        # Get webhook timestamp for correlation
        webhook_time = datetime.now()  # Approximate webhook receipt time
        
        analysis = {
            "webhook_notification": {  # Use the key the GUI expects
                "resource": resource,
                "changeType": change_type,  # Use camelCase to match GUI expectations
                "subscriptionId": subscription_id,  # Also use camelCase for consistency
                "webhook_timestamp": webhook_time.isoformat(),
                "correlation_window": "30 seconds"  # Time window for correlating changes
            },
            "webhook_info": {  # Keep this for backward compatibility
                "resource": resource,
                "changeType": change_type,
                "subscription_id": subscription_id,
                "webhook_timestamp": webhook_time.isoformat(),
                "correlation_window": "30 seconds"
            },
            "analysis_summary": [],
            "correlation_strategies": []
        }
        
        self.logger.info(f"Analyzing notification for resource: {resource}")
        self.logger.info(f"Change type: {change_type}")
        
        # Strategy 1: Get current item details (immediate state)
        analysis["correlation_strategies"].append("Strategy 1: Current item state analysis")
        item_details = self.get_item_details(resource)
        if item_details:
            # Save essential item information AND some useful expanded data
            analysis["item_details"] = {
                "name": item_details.get('name', 'Unknown'),
                "type": item_details.get('@odata.type', 'Unknown type'),
                "lastModifiedDateTime": item_details.get('lastModifiedDateTime'),
                "lastModifiedBy": item_details.get('lastModifiedBy', {}),
                "size": item_details.get('size'),
                "webUrl": item_details.get('webUrl'),
                "createdDateTime": item_details.get('createdDateTime'),
                "id": item_details.get('id'),
                # Include limited versions and activities if available
                "recent_versions": item_details.get('versions', {}).get('value', [])[:3] if item_details.get('versions') else [],
                "recent_activities": item_details.get('activities', {}).get('value', [])[:5] if item_details.get('activities') else []
            }
            
            # Include permissions for security analysis
            if 'permissions' in item_details:
                analysis["permissions"] = item_details['permissions']
                
            # Analyze what specifically changed
            self._analyze_specific_changes(analysis, item_details, webhook_time)
            
            analysis["analysis_summary"].append({
                "category": "Current Item State",
                "details": f"Item: {item_details.get('name', 'Unknown')} ({item_details.get('@odata.type', 'Unknown type')})",
                "timestamp": item_details.get('lastModifiedDateTime'),
                "correlation_method": "current_state"
            })
        else:
            self.logger.warning(f"Failed to get item details for resource: {resource}")
            analysis["analysis_summary"].append({
                "category": "Current Item State",
                "details": "[ERROR] Could not retrieve item details - API call failed",
                "correlation_method": "failed_api_call"
            })
            
            # Check if the item was recently modified (within correlation window)
            # Note: This will be skipped since we don't have item_details
            analysis["analysis_summary"].append({
                "category": "Time Correlation",
                "details": "[WARNING] Cannot perform time correlation - no item details",
                "correlation_method": "failed_api_call"
            })
            
            # Skip further analysis since we don't have item details
            return analysis
            
        # Strategy 2: Time correlation analysis (only if we have item_details)
        if item_details:
            last_modified = item_details.get('lastModifiedDateTime')
            if last_modified:
                try:
                    # Parse ISO datetime
                    modified_time = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                    time_diff = abs((webhook_time - modified_time.replace(tzinfo=None)).total_seconds())
                    
                    analysis["correlation_strategies"].append(f"Strategy 2: Time correlation - {time_diff:.1f}s difference")
                    
                    if time_diff <= 30:  # Within 30 seconds
                        analysis["analysis_summary"].append({
                            "category": "Time Correlation",
                            "details": f"[OK] Recent modification detected ({time_diff:.1f}s ago)",
                            "is_security_related": True,
                            "correlation_strength": "HIGH",
                            "correlation_method": "time_based"
                        })
                    else:
                        analysis["analysis_summary"].append({
                            "category": "Time Correlation",
                            "details": f"[WARNING] Modification not recent ({time_diff:.1f}s ago)",
                            "correlation_strength": "LOW",
                            "correlation_method": "time_based"
                        })
                        
                except Exception as e:
                    self.logger.warning(f"Could not parse modification time: {e}")
        
        # Strategy 3: Get permissions for security analysis  
        analysis["correlation_strategies"].append("Strategy 3: Permission and sharing analysis")
        
        # Use permissions from item_details if available, otherwise fetch separately
        permissions = None
        if item_details and 'permissions' in item_details:
            # Handle both list format (expanded) and dict format (separate call)
            perm_data = item_details['permissions']
            if isinstance(perm_data, list):
                permissions = perm_data  # Direct list from expand
            elif isinstance(perm_data, dict) and 'value' in perm_data:
                permissions = perm_data['value']  # Dictionary with value key
            else:
                self.logger.warning(f"Unexpected permissions format: {type(perm_data)}")
                permissions = []
        else:
            permissions = self.get_item_permissions(resource)
            
        if permissions:
            # Save permissions but with useful analysis
            sharing_changes = []
            permission_types = []
            
            for perm in permissions:
                if perm.get('link'):
                    link_info = perm['link']
                    link_type = link_info.get('type', 'unknown')
                    link_scope = link_info.get('scope', 'unknown')
                    permission_types.append(f"{link_type} link ({link_scope})")
                    sharing_changes.append(perm)
                elif 'anyone' in perm.get('grantedTo', {}).get('user', {}).get('displayName', '').lower():
                    sharing_changes.append(perm)
                elif perm.get('grantedTo', {}).get('user'):
                    user_info = perm['grantedTo']['user']
                    permission_types.append(f"User: {user_info.get('displayName', 'Unknown')}")
            
            # Include permissions data for review but summarize key findings
            analysis["permissions_analysis"] = {
                "total_permissions": len(permissions),
                "sharing_links_count": len([p for p in permissions if p.get('link')]),
                "permission_types": permission_types,
                "full_permissions": permissions  # Include full data for detailed review
            }
            
            if sharing_changes:
                analysis["analysis_summary"].append({
                    "category": "Security Analysis",
                    "details": f"[SECURITY] Found {len(sharing_changes)} sharing permission(s)",
                    "is_security_related": True,
                    "correlation_strength": "HIGH",
                    "correlation_method": "permission_analysis"
                })
            
            analysis["analysis_summary"].append({
                "category": "Permissions Overview",
                "details": f"Total permissions: {len(permissions)}",
                "correlation_method": "permission_analysis"
            })
        
        # Strategy 4: Resource-specific analysis
        analysis["correlation_strategies"].append("Strategy 4: Resource path analysis")
        if '/root' in resource:
            analysis["analysis_summary"].append({
                "category": "Resource Type",
                "details": "[FOLDER] Root-level resource (high visibility)",
                "is_security_related": True,
                "correlation_method": "resource_analysis"
            })
        elif '/items/' in resource:
            analysis["analysis_summary"].append({
                "category": "Resource Type", 
                "details": "[FILE] Specific item targeted",
                "correlation_method": "resource_analysis"
            })
        
        # Strategy 5: Look for recent audit activities (if available)
        analysis["correlation_strategies"].append("Strategy 5: Activity timeline analysis")
        try:
            # Check if we already have activities from item_details
            activities_from_item = analysis.get("item_details", {}).get("recent_activities", [])
            
            if activities_from_item:
                # Use activities from item expansion
                activity_analysis = {
                    "total_activities": len(activities_from_item),
                    "activities": activities_from_item,
                    "activity_types": list(set([act.get('action', {}).get('@odata.type', 'Unknown') for act in activities_from_item if act.get('action')]))
                }
                analysis["activity_analysis"] = activity_analysis
                
                analysis["analysis_summary"].append({
                    "category": "Recent Activities",
                    "details": f"[TIME] Found {len(activities_from_item)} recent activities from item",
                    "is_security_related": True,
                    "correlation_strength": "HIGH",
                    "correlation_method": "activity_timeline"
                })
            else:
                # For root folder changes, try to get recent drive-level activities
                if '/root' in resource and '/drives/' in resource:
                    drive_id = resource.split('/drives/')[1].split('/')[0]
                    recent_items = self._get_recent_drive_changes(drive_id)
                    
                    if recent_items:
                        analysis["recent_drive_changes"] = recent_items
                        analysis["analysis_summary"].append({
                            "category": "Recent Drive Changes",
                            "details": f"[DRIVE] Found {len(recent_items)} recent changes in drive: {recent_items[0]['name']}",
                            "is_security_related": True,
                            "correlation_strength": "HIGH",
                            "correlation_method": "drive_analysis"
                        })
                    else:
                        # Fall back to regular activity search
                        activities = self.get_recent_activities(resource, minutes=2)
                        if activities:
                            activity_analysis = {
                                "total_activities": len(activities),
                                "activities": activities[:5],
                                "activity_names": [activity.get('name', 'Unknown') for activity in activities[:5]]
                            }
                            analysis["activity_analysis"] = activity_analysis
                            
                            analysis["analysis_summary"].append({
                                "category": "Recent Activities",
                                "details": f"[TIME] Found {len(activities)} recent activities from drive",
                                "is_security_related": True,
                                "correlation_strength": "MEDIUM",
                                "correlation_method": "activity_timeline"
                            })
                        else:
                            analysis["analysis_summary"].append({
                                "category": "Activity Timeline",
                                "details": "[INFO] No recent activities found - this may be a folder metadata change",
                                "correlation_method": "activity_timeline"
                            })
                else:
                    analysis["analysis_summary"].append({
                        "category": "Activity Timeline", 
                        "details": "[INFO] No recent activities found",
                        "correlation_method": "activity_timeline"
                    })
                    
        except Exception as e:
            self.logger.warning(f"Could not retrieve recent activities: {e}")
            analysis["analysis_summary"].append({
                "category": "Activity Timeline",
                "details": "[WARNING] Could not retrieve recent activities",
                "correlation_method": "activity_timeline"
            })
        
        return analysis

    def analyze_change_details(self, notification: Dict[str, Any]) -> Dict[str, Any]:
        """
        Focus on finding actual changes - what files were added, modified, or removed.
        
        When a webhook is for a folder, the actual change is usually in the folder's contents.
        """
        self.logger.info("Analyzing webhook to find actual file/content changes")
        
        resource = notification.get('resource', '')
        change_type = notification.get('changeType', '')
        subscription_id = notification.get('subscriptionId', '')
        
        analysis_time = datetime.now()
        
        # Try to get the actual webhook timestamp for more accurate correlation
        webhook_timestamp = getattr(self, '_current_webhook_timestamp', None)
        if webhook_timestamp:
            try:
                # Webhook timestamp is in local time (EST), convert to UTC
                local_time = datetime.fromisoformat(webhook_timestamp.replace('Z', ''))
                analysis_time = local_time + timedelta(hours=4)  # EDT to UTC (until Nov 2nd)
                self.logger.info(f"Webhook {webhook_timestamp} (EST) -> {analysis_time} UTC")
            except Exception as e:
                self.logger.warning(f"Could not parse webhook timestamp: {e}")
                analysis_time = datetime.now()
        
        # Initialize analysis focused on finding actual changes
        analysis = {
            "webhook_notification": {
                "resource": resource,
                "changeType": change_type,
                "subscriptionId": subscription_id,
                "timestamp": analysis_time.isoformat(),
                "approach": "find_actual_changes"
            },
            "analysis_summary": []
        }
        
        self.logger.info(f"Looking for actual changes in: {resource}")
        self.logger.info(f"Change type: {change_type}")
        
        # For folder notifications, focus on what changed IN the folder
        if '/root' in resource or resource.endswith('/root'):
            # This is a root folder change - look for recently added/modified files
            analysis["analysis_summary"].append({
                "category": "Change Location",
                "details": "Root folder notification - checking for file changes within folder",
                "correlation_method": "folder_analysis"
            })
            
            # Get recent items in the folder to see what actually changed
            try:
                children_url = f"https://graph.microsoft.com/v1.0/{resource.lstrip('/')}/children?$top=20&$orderby=lastModifiedDateTime desc"
                children_response = self._make_graph_request(children_url)
                
                if children_response and 'value' in children_response:
                    recent_items = children_response['value']
                    
                    # Look for changes that happened before the webhook (within last 2 hours)
                    cutoff_time = analysis_time - timedelta(hours=2)
                    self.logger.info(f"Looking for changes between {cutoff_time} and {analysis_time}")
                    recent_changes = []
                    
                    for item in recent_items:
                        try:
                            item_id = item.get('id')
                            item_name = item.get('name', 'Unknown')
                            last_modified = item.get('lastModifiedDateTime')
                            
                            self.logger.info(f"Checking item: {item_name}, modified: {last_modified}")
                            
                            # Check if this item had recent file modifications
                            file_recently_modified = False
                            if last_modified:
                                modified_time = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                                if modified_time.replace(tzinfo=None) > cutoff_time:
                                    file_recently_modified = True
                                    self.logger.info(f"File {item_name} was recently modified")
                            
                            # ALSO check for recent permission/sharing activities on this specific file
                            permission_activities = []
                            if item_id:
                                try:
                                    # Get activities for this specific item - REMOVE /root from the URL
                                    drive_id = resource.split('/')[2]  # Extract drive ID from resource
                                    activities_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/activities?$top=10"
                                    activities_response = self._make_graph_request(activities_url)
                                    
                                    if activities_response and 'value' in activities_response:
                                        for activity in activities_response['value']:
                                            activity_time_str = activity.get('times', {}).get('recordedDateTime')
                                            if activity_time_str:
                                                try:
                                                    activity_time = datetime.fromisoformat(activity_time_str.replace('Z', '+00:00'))
                                                    if activity_time.replace(tzinfo=None) > cutoff_time:
                                                        # Recent activity on this file!
                                                        action = activity.get('action', {})
                                                        actor = activity.get('actor', {})
                                                        
                                                        # Determine the specific type of operation
                                                        operation_type = 'unknown'
                                                        operation_details = ''
                                                        
                                                        if 'share' in action:
                                                            # Permission change - check if it's add or remove
                                                            share_action = action['share']
                                                            recipients = share_action.get('recipients', [])
                                                            
                                                            # Check if this is a permission grant or revoke
                                                            # The presence of recipients usually indicates a grant
                                                            if recipients:
                                                                operation_type = 'permission_granted'
                                                                user_names = []
                                                                for recipient in recipients:
                                                                    user_info = recipient.get('user', {})
                                                                    user_name = user_info.get('displayName', user_info.get('email', 'Unknown'))
                                                                    user_names.append(user_name)
                                                                operation_details = f"Access granted to: {', '.join(user_names)}"
                                                            else:
                                                                # No recipients might indicate a permission removal or link sharing
                                                                operation_type = 'permission_modified'
                                                                operation_details = "Permission settings modified"
                                                        
                                                        elif 'create' in action:
                                                            operation_type = 'file_uploaded'
                                                            operation_details = "New file uploaded to folder"
                                                        
                                                        elif 'edit' in action:
                                                            # Check if it's a content edit or metadata change
                                                            if 'version' in action:
                                                                operation_type = 'file_content_modified'
                                                                version_info = action.get('version', {})
                                                                new_version = version_info.get('newVersion', 'unknown')
                                                                operation_details = f"File content edited (version {new_version})"
                                                            else:
                                                                operation_type = 'file_modified'
                                                                operation_details = "File metadata or content modified"
                                                        
                                                        elif 'rename' in action:
                                                            operation_type = 'file_renamed'
                                                            rename_info = action.get('rename', {})
                                                            old_name = rename_info.get('oldName', 'Unknown')
                                                            operation_details = f"File renamed from: {old_name}"
                                                        
                                                        elif 'move' in action:
                                                            operation_type = 'file_moved'
                                                            move_info = action.get('move', {})
                                                            from_location = move_info.get('from', 'Unknown location')
                                                            operation_details = f"File moved from: {from_location}"
                                                        
                                                        elif 'delete' in action:
                                                            operation_type = 'file_deleted'
                                                            operation_details = "File deleted"
                                                        
                                                        elif 'restore' in action:
                                                            operation_type = 'file_restored'
                                                            operation_details = "File restored from deletion"
                                                        
                                                        elif 'copy' in action:
                                                            operation_type = 'file_copied'
                                                            operation_details = "File copied"
                                                        
                                                        else:
                                                            # Fallback for unknown actions
                                                            operation_type = 'unknown_operation'
                                                            action_keys = list(action.keys())
                                                            operation_details = f"Action type: {', '.join(action_keys)}"
                                                        
                                                        permission_activities.append({
                                                            'type': operation_type,
                                                            'time': activity_time_str,
                                                            'action': action,
                                                            'actor': actor,
                                                            'operation_details': operation_details
                                                        })
                                                except Exception:
                                                    continue
                                except Exception as e:
                                    self.logger.debug(f"Could not get activities for {item_name}: {e}")
                            
                            # Add to recent changes if either file was modified OR had permission changes
                            if file_recently_modified or permission_activities:
                                change_entry = {
                                    'name': item_name,
                                    'type': 'folder' if item.get('folder') else 'file',
                                    'size': item.get('size'),
                                    'modified': last_modified,
                                    'modified_by': item.get('lastModifiedBy', {}),
                                    'created': item.get('createdDateTime'),
                                    'id': item_id,
                                    'web_url': item.get('webUrl'),
                                    'file_recently_modified': file_recently_modified,
                                    'permission_activities': permission_activities
                                }
                                recent_changes.append(change_entry)
                                
                        except Exception as e:
                            # Skip items with parsing errors
                            self.logger.debug(f"Could not parse item modification time: {e}")
                            continue
                    
                    if recent_changes:
                        # Find the change that most closely correlates with the webhook timestamp
                        webhook_time = analysis_time
                        best_match = None
                        best_match_score = float('inf')
                        
                        for change in recent_changes:
                            # Check both file modification time and permission activity times
                            change_times = []
                            
                            # Add file modification time if recent
                            if change.get('file_recently_modified') and change.get('modified'):
                                try:
                                    mod_time = datetime.fromisoformat(change['modified'].replace('Z', '+00:00'))
                                    change_times.append(('file_modification', mod_time.replace(tzinfo=None), change['modified'], ''))
                                except Exception:
                                    pass
                            
                            # Add permission activity times with operation type preference
                            for perm_activity in change.get('permission_activities', []):
                                try:
                                    perm_time = datetime.fromisoformat(perm_activity['time'].replace('Z', '+00:00'))
                                    operation_type = perm_activity.get('type', 'unknown_operation')
                                    operation_details = perm_activity.get('operation_details', '')
                                    change_times.append((operation_type, perm_time.replace(tzinfo=None), perm_activity['time'], operation_details))
                                except Exception:
                                    pass
                            
                            # Find the time closest to the webhook (activities should be BEFORE webhook)
                            for change_type, change_time, time_str, *extra_details in change_times:
                                operation_details = extra_details[0] if extra_details else ''
                                # Calculate if this change happened before the webhook
                                time_diff = (webhook_time - change_time).total_seconds()  # Webhook - Change
                                
                                # We want changes that happened before the webhook (positive time_diff)
                                # but not too long ago (within reasonable webhook delay)
                                if 0 < time_diff < 7200:  # Change happened before webhook, within 2 hours
                                    score = time_diff
                                    
                                    # Apply preference scoring based on operation type
                                    if change_type in ['permission_granted', 'permission_modified']:
                                        score = score * 0.6  # Highest preference for permission changes
                                    elif change_type in ['file_uploaded', 'file_renamed', 'file_moved']:
                                        score = score * 0.8  # High preference for significant file operations
                                    elif change_type in ['file_content_modified', 'file_modified']:
                                        score = score * 0.9  # Medium preference for file modifications
                                    elif change_type == 'file_modification':
                                        score = score * 1.0  # Standard preference for basic file changes
                                    
                                    if score < best_match_score:
                                        best_match_score = score
                                        best_match = {
                                            'change': change,
                                            'change_type': change_type,
                                            'change_time': change_time,
                                            'time_str': time_str,
                                            'operation_details': operation_details,
                                            'latency_seconds': time_diff  # How long after change the webhook arrived
                                        }
                        
                        if best_match and best_match['latency_seconds'] < 7200:  # Within 2 hours (more generous)
                            change = best_match['change']
                            change_type = best_match['change_type']
                            latency = best_match['latency_seconds']
                            
                            analysis["webhook_correlation"] = {
                                "matched_item": change['name'],
                                "change_type": change_type,
                                "operation_details": best_match.get('operation_details', ''),
                                "change_time": best_match['time_str'],
                                "webhook_time": webhook_time.isoformat(),
                                "latency_seconds": latency,
                                "correlation_confidence": "high" if latency < 300 else "medium" if latency < 1800 else "low"
                            }
                            
                            if change_type == 'permission_change':
                                # Extract permission details
                                perm_activities = change.get('permission_activities', [])
                                recipients = []
                                for perm_activity in perm_activities:
                                    share_action = perm_activity.get('action', {}).get('share', {})
                                    for recipient in share_action.get('recipients', []):
                                        user_info = recipient.get('user', {})
                                        user_name = user_info.get('displayName', user_info.get('email', 'Unknown'))
                                        recipients.append(user_name)
                                
                                analysis["analysis_summary"].append({
                                    "category": "WEBHOOK CORRELATION",
                                    "details": f"Permission change detected on '{change['name']}' - shared with: {', '.join(recipients)}",
                                    "correlation_method": "timing_analysis"
                                })
                                
                                analysis["analysis_summary"].append({
                                    "category": "TIMING ANALYSIS", 
                                    "details": f"Permission change at {best_match['time_str']}, webhook received {latency:.0f} seconds later",
                                    "correlation_method": "timing_analysis"
                                })
                                
                                analysis["analysis_summary"].append({
                                    "category": "CHANGE DETAILS",
                                    "details": f"File: {change['name']} ({change.get('size', 'unknown')} bytes) - Permission granted to {len(recipients)} user(s)",
                                    "correlation_method": "permission_analysis"
                                })
                                
                            else:  # file_modification
                                analysis["analysis_summary"].append({
                                    "category": "WEBHOOK CORRELATION",
                                    "details": f"File modification detected on '{change['name']}'",
                                    "correlation_method": "timing_analysis"
                                })
                                
                                analysis["analysis_summary"].append({
                                    "category": "TIMING ANALYSIS",
                                    "details": f"File modified at {best_match['time_str']}, webhook received {latency:.0f} seconds later", 
                                    "correlation_method": "timing_analysis"
                                })
                                
                                size_info = f" ({change.get('size', 'unknown')} bytes)" if change.get('size') else ""
                                modified_by = change.get('modified_by', {}).get('user', {}).get('displayName', 'Unknown')
                                analysis["analysis_summary"].append({
                                    "category": "📝 CHANGE DETAILS",
                                    "details": f"File: {change['name']}{size_info} - Modified by {modified_by}",
                                    "correlation_method": "file_analysis"
                                })
                        else:
                            analysis["analysis_summary"].append({
                                "category": "❓ WEBHOOK CORRELATION",
                                "details": "No recent changes found that closely correlate with this webhook timestamp",
                                "correlation_method": "timing_analysis"
                            })
                            
                            # Show the most recent item anyway for context
                            if recent_changes:
                                latest_item = recent_changes[0]
                                analysis["analysis_summary"].append({
                                    "category": "LATEST ACTIVITY",
                                    "details": f"Most recent: {latest_item.get('name', 'Unknown')} - Last activity: {latest_item.get('modified', 'Unknown')}",
                                    "correlation_method": "recent_items_analysis"
                                })
                    else:
                        analysis["analysis_summary"].append({
                            "category": "FOLDER ANALYSIS",
                            "details": "No recent changes found in folder (within 60 minutes)",
                            "correlation_method": "recent_items_analysis"
                        })
                
                else:
                    analysis["analysis_summary"].append({
                        "category": "Folder Contents",
                        "details": "Could not retrieve folder contents",
                        "correlation_method": "failed_api_call"
                    })
                    
            except Exception as e:
                analysis["analysis_summary"].append({
                    "category": "Folder Analysis",
                    "details": f"Error analyzing folder contents: {str(e)[:100]}",
                    "correlation_method": "failed_api_call"
                })
        
        else:
            # This is a specific file change
            item_details = self.get_item_details(resource)
            if item_details:
                analysis["item_details"] = item_details
                
                analysis["analysis_summary"].append({
                    "category": "Direct File Change",
                    "details": f"File: {item_details.get('name', 'Unknown')} was {change_type}",
                    "correlation_method": "direct_item_analysis"
                })
                
                if item_details.get('lastModifiedDateTime'):
                    analysis["analysis_summary"].append({
                        "category": "File Details",
                        "details": f"Size: {item_details.get('size', 'Unknown')} bytes, Modified: {item_details['lastModifiedDateTime']}",
                        "correlation_method": "direct_item_analysis"
                    })
        
        # Add general webhook info (but don't add redundant timing if we already found a correlation)
        if not any(item.get('category', '').startswith('TIMING') for item in analysis.get("analysis_summary", [])):
            analysis["analysis_summary"].append({
                "category": "WEBHOOK INFO",
                "details": f"Webhook received for {change_type} change on folder at {analysis_time.strftime('%H:%M:%S UTC')}",
                "correlation_method": "webhook_metadata"
            })
            
        return analysis
    
    def process_webhook_notification(self, notification: Dict[str, Any]) -> Dict[str, Any]:
        """Process a webhook notification and get detailed analysis"""
        try:
            analysis = self.analyze_change_details(notification)
            
            # Save detailed analysis
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Ensure change_analysis directory exists
            import os
            os.makedirs("change_analysis", exist_ok=True)
            
            analysis_file = f"change_analysis/enhanced_analysis_{timestamp}.json"
            
            with open(analysis_file, 'w', encoding='utf-8') as f:
                json.dump(analysis, f, indent=2)
            
            self.logger.info(f"Enhanced analysis saved to: {analysis_file}")
            
            # Log summary
            summary_count = len(analysis.get('analysis_summary', []))
            self.logger.info(f"Analysis complete: {summary_count} insights found")
            
            for insight in analysis.get('analysis_summary', []):
                security_note = " [SECURITY-RELATED]" if insight.get('is_security_related', False) else ""
                self.logger.info(f"  {insight.get('category', 'Unknown')}: {insight.get('details', 'N/A')}{security_note}")
            
            return analysis
            
        except Exception as e:
            self.logger.error(f"Error processing webhook notification: {e}")
            return {"error": str(e)}

def analyze_webhook_notification_file(file_path: str) -> Dict[str, Any]:
    """Analyze a webhook notification file using enhanced tracking"""
    try:
        # Load webhook notification
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        notification_data = data.get("notification", {})
        if "value" not in notification_data:
            print(f"ERROR: No notifications found in {file_path}")
            return {}
        
        # Initialize enhanced tracker
        tracker = EnhancedChangeTracker()
        
        notifications = notification_data["value"]
        
        print(f"INFO: Analyzing {len(notifications)} notification(s) from {file_path}")
        
        all_analyses = []
        for i, notification in enumerate(notifications, 1):
            print(f"\n--- Analyzing Notification {i} ---")
            analysis = tracker.process_webhook_notification(notification)
            all_analyses.append(analysis)
        
        return {"analyses": all_analyses, "total_notifications": len(notifications)}
        
    except Exception as e:
        print(f"ERROR: Error analyzing webhook notification file: {e}")
        return {"error": str(e)}

def main():
    """Main function for testing"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Enhanced Change Tracker for Microsoft Graph")
    parser.add_argument("--webhook-file", type=str, help="Webhook notification file to analyze")
    
    args = parser.parse_args()
    
    if args.webhook_file:
        # Analyze a specific webhook notification file
        result = analyze_webhook_notification_file(args.webhook_file)
        
        if "error" in result:
            print(f"ERROR: Analysis failed: {result['error']}")
        elif "analyses" in result:
            print(f"\n[COMPLETE] Analysis complete for {result['total_notifications']} notification(s)")
            
            for i, analysis in enumerate(result["analyses"], 1):
                if "error" in analysis:
                    print(f"  Notification {i}: Failed - {analysis['error']}")
                else:
                    summary_count = len(analysis.get('analysis_summary', []))
                    print(f"  Notification {i}: {summary_count} insights found")
    else:
        print("Usage: python enhanced_change_tracker.py --webhook-file <file>")

if __name__ == "__main__":
    main()