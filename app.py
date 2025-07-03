# ─────────────────────────────────────────────────────────────
# SIMPLE DIAGNOSTIC SCRIPT - RUN THIS FIRST
# ─────────────────────────────────────────────────────────────

import streamlit as st
import requests
import os

st.title("🔍 SharePoint Connection Diagnostic")

# Get your configuration
SITE_URL = os.getenv("SP_SITE_URL") or st.secrets.get("SP_SITE_URL", "")
USERNAME = os.getenv("SP_USERNAME") or st.secrets.get("SP_USERNAME", "")

st.write("**Your Configuration:**")
st.write(f"Site URL: {SITE_URL}")
st.write(f"Username: {USERNAME}")

# Test 1: Can we reach SharePoint?
st.write("---")
st.subheader("Test 1: Basic Connectivity")

if st.button("Test Site Connectivity"):
    try:
        response = requests.get(SITE_URL, timeout=10)
        st.success(f"✅ Site is reachable (Status: {response.status_code})")
        
        # Check authentication requirements
        auth_url = f"{SITE_URL}/_api/web"
        auth_response = requests.get(auth_url, timeout=10)
        
        if auth_response.status_code == 401:
            auth_header = auth_response.headers.get('www-authenticate', '')
            st.info(f"🔐 Authentication required: {auth_header}")
            
            if 'Bearer' in auth_header:
                st.error("❌ **Modern Authentication Required**")
                st.error("Your organization requires OAuth/Modern Authentication")
                st.error("Basic username/password won't work")
            else:
                st.info("ℹ️ Basic authentication might be supported")
        
    except Exception as e:
        st.error(f"❌ Cannot reach site: {str(e)}")

# Test 2: Check account type
st.write("---")
st.subheader("Test 2: Account Analysis")

if USERNAME:
    st.write("**Account Type Analysis:**")
    
    if "@" in USERNAME:
        st.success("✅ Username is in email format")
        domain = USERNAME.split("@")[1]
        st.write(f"Domain: {domain}")
        
        if ".onmicrosoft.com" in domain:
            st.info("ℹ️ This is a cloud-only account")
        else:
            st.info("ℹ️ This might be a federated/hybrid account")
    else:
        st.warning("⚠️ Username should be in email format")
    
    # Common issues checklist
    st.write("**Common Issues Checklist:**")
    
    issues = [
        "✅ Is Multi-Factor Authentication (MFA) enabled?",
        "✅ Has the password been changed recently?",
        "✅ Is the account locked or disabled?",
        "✅ Are there Conditional Access policies?",
        "✅ Does the account have SharePoint permissions?",
        "✅ Is this a personal OneDrive site (needs different permissions)?",
    ]
    
    for issue in issues:
        st.write(issue)

# Test 3: What to do next
st.write("---")
st.subheader("Test 3: Next Steps")

st.write("**If you're getting AADSTS80002 error:**")

col1, col2 = st.columns(2)

with col1:
    st.write("**🔥 Quick Fixes to Try:**")
    st.write("1. **Check MFA**: If MFA is enabled, you need App-Only auth")
    st.write("2. **Try different account**: Use account without MFA")
    st.write("3. **Check password**: Verify it's correct and recent")
    st.write("4. **Test in browser**: Can you access SharePoint manually?")

with col2:
    st.write("**🔧 Technical Solutions:**")
    st.write("1. **App-Only Authentication** (recommended)")
    st.write("2. **Service Account** without MFA")
    st.write("3. **REST API approach** (bypasses office365 library)")
    st.write("4. **Different authentication library**")

# Test 4: Quick App-Only Setup Check
st.write("---")
st.subheader("Test 4: App-Only Authentication Setup")

CLIENT_ID = os.getenv("SP_CLIENT_ID") or st.secrets.get("SP_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET") or st.secrets.get("SP_CLIENT_SECRET", "")

if CLIENT_ID and CLIENT_SECRET:
    st.success("✅ App-Only credentials found!")
    st.write(f"Client ID: {CLIENT_ID}")
    st.write("Client Secret: ***configured***")
    
    # Try app-only authentication
    if st.button("Test App-Only Authentication"):
        try:
            from office365.sharepoint.client_context import ClientContext
            from office365.runtime.auth.client_credential import ClientCredential
            
            credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
            ctx = ClientContext(SITE_URL).with_credentials(credentials)
            
            # Test connection
            ctx.load(ctx.web)
            ctx.execute_query()
            
            st.success("✅ App-Only Authentication works!")
            st.success("Use this method in your main app")
            
        except Exception as e:
            st.error(f"❌ App-Only failed: {str(e)}")
            
            if "AADSTS70011" in str(e):
                st.error("❌ Invalid scope. Check your API permissions.")
            elif "AADSTS700016" in str(e):
                st.error("❌ Application not found. Check your Client ID.")
            elif "AADSTS7000215" in str(e):
                st.error("❌ Invalid client secret. Check your Client Secret.")
else:
    st.warning("⚠️ App-Only credentials not configured")
    st.info("👉 This is likely your solution - set up App-Only authentication")
    
    if st.button("Show App-Only Setup Instructions"):
        st.markdown("""
        ### Quick Setup for App-Only Authentication:
        
        1. **Go to Azure Portal**: https://portal.azure.com
        2. **Azure Active Directory** → **App registrations** → **New registration**
        3. **Name**: "Dismac SharePoint App"
        4. **Register** the app
        5. **API permissions** → **Add permission** → **SharePoint** → **Application permissions** → **Sites.ReadWrite.All**
        6. **Grant admin consent**
        7. **Certificates & secrets** → **New client secret** → **Copy the value**
        8. **Add to your secrets**:
        ```toml
        SP_CLIENT_ID = "your-app-id"
        SP_CLIENT_SECRET = "your-secret"
        SP_TENANT_ID = "your-tenant-id"
        ```
        """)

# Final recommendations
st.write("---")
st.subheader("🎯 Recommended Solution")

st.success("**Most likely solution for AADSTS80002:**")
st.success("1. Set up App-Only Authentication (steps above)")
st.success("2. Use ClientCredential instead of UserCredential")
st.success("3. This bypasses MFA and modern auth requirements")

st.info("**Alternative if you can't set up App-Only:**")
st.info("1. Create a service account without MFA")
st.info("2. Give it SharePoint permissions")
st.info("3. Use that account for the application")