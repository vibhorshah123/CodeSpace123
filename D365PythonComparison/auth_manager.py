"""
Authentication Manager for Dynamics 365
Handles authentication using different methods (browser-based OAuth or client credentials)
"""

import requests
import webbrowser
import secrets
import hashlib
import base64
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, urlencode
from typing import Dict, Optional
import threading


class AuthManager:
    """Manages authentication and token acquisition for D365 environments"""
    
    def __init__(self, credentials: Dict[str, str]):
        """
        Initialize authentication manager
        
        Args:
            credentials: Dictionary containing authentication details
        """
        self.credentials = credentials
        self.auth_type = credentials.get("auth_type")
        self._cached_tokens: Dict[str, str] = {}
        self._device_code_used = False
    
    def get_token(self, environment_url: str) -> str:
        """
        Get access token for the specified environment
        
        Args:
            environment_url: URL of the D365 environment
            
        Returns:
            Access token string
            
        Raises:
            Exception: If authentication fails
        """
        # Check cache first
        if environment_url in self._cached_tokens:
            return self._cached_tokens[environment_url]
        
        # Acquire new token based on auth type
        if self.auth_type == "browser_login":
            token = self._get_token_browser_oauth(environment_url)
        elif self.auth_type == "client_credentials":
            token = self._get_token_client_credentials(environment_url)
        else:
            raise ValueError(f"Unsupported auth type: {self.auth_type}")
        
        # Cache the token
        self._cached_tokens[environment_url] = token
        return token
    
    def _get_token_browser_oauth(self, environment_url: str) -> str:
        """
        Acquire token using OAuth authorization code flow with local redirect
        Opens browser for user to login and captures the token via local server
        
        Args:
            environment_url: URL of the D365 environment
            
        Returns:
            Access token
        """
        # Extract the resource URL
        parsed = urlparse(environment_url)
        scope = f"{parsed.scheme}://{parsed.netloc}/.default"
        
        # OAuth configuration
        tenant = "common"
        client_id = "51f81489-12ee-4a9e-aaae-a2591f45987d"  # Microsoft Dynamics CRM client ID
        redirect_uri = "http://localhost:8400"
        
        # Generate PKCE values
        code_verifier = secrets.token_urlsafe(64)
        code_challenge = base64.urlsafe_b64encode(
            hashlib.sha256(code_verifier.encode()).digest()
        ).decode().rstrip('=')
        
        # Authorization URL
        auth_params = {
            "client_id": client_id,
            "response_type": "code",
            "redirect_uri": redirect_uri,
            "response_mode": "query",
            "scope": scope,
            "state": secrets.token_urlsafe(16),
            "code_challenge": code_challenge,
            "code_challenge_method": "S256",
            "prompt": "select_account"
        }
        
        auth_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?{urlencode(auth_params)}"
        
        # Storage for the authorization code
        auth_data = {"code": None, "error": None}
        
        class OAuthCallbackHandler(BaseHTTPRequestHandler):
            def log_message(self, format, *args):
                pass  # Suppress logging
            
            def do_GET(self):
                # Parse the callback URL
                query_components = parse_qs(urlparse(self.path).query)
                
                if 'code' in query_components:
                    auth_data['code'] = query_components['code'][0]
                    # Send success page
                    self.send_response(200)
                    self.send_header('Content-type', 'text/html')
                    self.end_headers()
                    success_html = """
                    <html>
                    <head><title>Authentication Successful</title></head>
                    <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
                        <h1 style="color: #28a745;">✓ Authentication Successful!</h1>
                        <p>You have successfully signed in to Dynamics 365.</p>
                        <p><strong>You can close this window and return to the application.</strong></p>
                    </body>
                    </html>
                    """
                    self.wfile.write(success_html.encode())
                elif 'error' in query_components:
                    auth_data['error'] = query_components.get('error_description', ['Unknown error'])[0]
                    # Send error page
                    self.send_response(200)
                    self.send_header('Content-type', 'text/html')
                    self.end_headers()
                    error_html = f"""
                    <html>
                    <head><title>Authentication Failed</title></head>
                    <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
                        <h1 style="color: #dc3545;">✗ Authentication Failed</h1>
                        <p>{auth_data['error']}</p>
                        <p>Please close this window and try again.</p>
                    </body>
                    </html>
                    """
                    self.wfile.write(error_html.encode())
        
        # Start local server
        server = HTTPServer(('localhost', 8400), OAuthCallbackHandler)
        server.timeout = 300  # 5 minutes timeout
        
        # Open browser
        print("\n" + "=" * 70)
        print("  AUTHENTICATION REQUIRED")
        print("=" * 70)
        print("\n  Opening browser for authentication...")
        print("  Please sign in with your Microsoft account.")
        print("\n  If browser doesn't open automatically, visit:")
        print(f"  {auth_url[:80]}...")
        print("\n" + "=" * 70)
        print()
        
        webbrowser.open(auth_url)
        
        # Wait for callback
        print("  Waiting for authentication...")
        server.handle_request()
        server.server_close()
        
        if auth_data['error']:
            raise Exception(f"Authentication failed: {auth_data['error']}")
        
        if not auth_data['code']:
            raise Exception("No authorization code received")
        
        print("  Authorization code received, exchanging for token...")
        
        # Exchange authorization code for token
        token_endpoint = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
        
        token_data = {
            "client_id": client_id,
            "grant_type": "authorization_code",
            "code": auth_data['code'],
            "redirect_uri": redirect_uri,
            "code_verifier": code_verifier,
            "scope": scope
        }
        
        try:
            response = requests.post(token_endpoint, data=token_data, timeout=30)
            response.raise_for_status()
            
            result = response.json()
            
            if "access_token" not in result:
                raise Exception("No access token in response")
            
            print("  ✓ Authentication successful!\n")
            return result["access_token"]
            
        except requests.RequestException as e:
            raise Exception(f"Failed to exchange code for token: {str(e)}")
    
    def _get_token_client_credentials(self, environment_url: str) -> str:
        """
        Acquire token using client credentials (application authentication)
        
        Args:
            environment_url: URL of the D365 environment
            
        Returns:
            Access token
        """
        tenant_id = self.credentials["tenant_id"]
        token_endpoint = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        
        # Extract the resource URL
        parsed = urlparse(environment_url)
        scope = f"{parsed.scheme}://{parsed.netloc}/.default"
        
        # Prepare request
        data = {
            "grant_type": "client_credentials",
            "client_id": self.credentials["client_id"],
            "client_secret": self.credentials["client_secret"],
            "scope": scope
        }
        
        try:
            response = requests.post(token_endpoint, data=data, timeout=30)
            response.raise_for_status()
            
            result = response.json()
            
            if "access_token" not in result:
                raise Exception("No access token in response")
            
            return result["access_token"]
            
        except requests.RequestException as e:
            raise Exception(f"Failed to acquire token: {str(e)}")
    
    def clear_cache(self):
        """Clear cached tokens"""
        self._cached_tokens.clear()
