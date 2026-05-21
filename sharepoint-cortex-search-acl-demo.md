# SharePoint to Cortex Search with Document ACLs — End-to-End Demo

> **Deployment type:** SPCS (Snowflake-managed runtime)  
> **Connector variant:** `unstructured-sharepoint-cdc` (Cortex Search + ACLs)  
> **Account:** ZIB31868 | **Runtime:** `spcs1runtime1`

---

## Architecture Overview

```
┌──────────────────┐       ┌─────────────────────────┐       ┌──────────────────────────┐
│  SharePoint Site │──────▶│  Openflow SPCS Runtime  │──────▶│  Snowflake Cortex Search │
│  (Documents +    │       │  (NiFi-based connector) │       │  (with ACL filtering)    │
│   Permissions)   │       └─────────────────────────┘       └──────────────────────────┘
└──────────────────┘                                                     │
                                                                         ▼
                                                              ┌──────────────────────┐
                                                              │  Query with user_email│
                                                              │  filter for ACL       │
                                                              └──────────────────────┘
```

**What this does:**
1. Ingests documents from a SharePoint site into Snowflake
2. Parses documents using `AI_PARSE_DOCUMENT`
3. Chunks text and creates a Cortex Search Service
4. Syncs document-level ACLs (user IDs and emails) from SharePoint
5. Enables ACL-filtered queries so users only see documents they have access to

---

## Prerequisites

### 1. Microsoft Entra ID (Azure AD) App Registration

| Item | How to Obtain |
|------|---------------|
| **Tenant ID** | Entra Admin Center → Overview → Tenant ID |
| **Client ID** | App Registration → Application (client) ID |
| **Client Secret** | App Registration → Certificates & secrets → New client secret |
| **Certificate (PEM)** | Generate or upload (required for ACL resolution) |
| **Private Key (PEM)** | Corresponding unencrypted private key for the certificate |

#### Required API Permissions (Application type, NOT delegated)

| Permission | Purpose |
|------------|---------|
| `Sites.Selected` (Microsoft Graph) | Limits access to specified sites only |
| `GroupMember.Read.All` (Microsoft Graph) | Resolves SharePoint group permissions |
| `User.ReadBasic.All` (Microsoft Graph) | Resolves Microsoft 365 user emails |
| `Sites.Selected` (SharePoint) | Required in addition to Graph permission |

> **Admin consent required** for all permissions above.

#### Grant Site Access

Grant the `fullcontrol` role to the application on the target site using PowerShell (PnP):

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<sitename>" -Interactive

Grant-PnPAzureADAppSitePermission `
  -AppId "<client-id>" `
  -DisplayName "Openflow SharePoint Connector" `
  -Permissions FullControl `
  -Site "https://<tenant>.sharepoint.com/sites/<sitename>"
```

Or via Microsoft Graph API:

```bash
curl -X POST "https://graph.microsoft.com/v1.0/sites/<site-id>/permissions" \
  -H "Authorization: Bearer <admin-token>" \
  -H "Content-Type: application/json" \
  -d '{
    "roles": ["fullcontrol"],
    "grantedToIdentities": [{
      "application": {
        "id": "<client-id>",
        "displayName": "Openflow SharePoint Connector"
      }
    }]
  }'
```

#### Generate Certificate + Private Key (if needed)

```bash
openssl req -x509 -nodes -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 \
  -subj "/CN=OpenflowSharePointConnector"
```

Upload `cert.pem` to the App Registration under **Certificates & secrets → Certificates**.

### 2. SharePoint Site Information

| Item | Example Value |
|------|---------------|
| Site URL | `https://contoso.sharepoint.com/sites/Engineering` |
| Document Library Name | `Shared Documents` (default) or custom name |
| Source Folder | `/` for root, or `/subfolder/path` |
| File Extensions | `pdf,docx,pptx,xlsx` (or leave empty for all) |

### 3. Snowflake Prerequisites

Confirm the following exist or will be created:

| Item | Value |
|------|-------|
| Destination Database | e.g., `SHAREPOINT_DEMO` |
| Destination Schema | e.g., `CORTEX` |
| Warehouse | e.g., `SI_DEMO` |
| Runtime Role | `OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1` |

---

## Step 1: Create Snowflake Database and Schema

```sql
USE ROLE ACCOUNTADMIN;

CREATE DATABASE IF NOT EXISTS SHAREPOINT_DEMO;
CREATE SCHEMA IF NOT EXISTS SHAREPOINT_DEMO.CORTEX;

-- Grant permissions to the Openflow runtime role
GRANT USAGE ON DATABASE SHAREPOINT_DEMO TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT USAGE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE TABLE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE DYNAMIC TABLE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE STAGE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE SEQUENCE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE CORTEX SEARCH SERVICE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;

-- Grant warehouse access
GRANT USAGE, OPERATE ON WAREHOUSE SI_DEMO TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
```

---

## Step 2: Configure External Access Integration (EAI)

SPCS runtimes cannot reach external endpoints without an EAI. The SharePoint connector requires:

| Domain | Purpose |
|--------|---------|
| `login.microsoftonline.com:443` | OAuth authentication |
| `login.microsoft.com:443` | OAuth authentication (fallback) |
| `graph.microsoft.com:443` | ACL resolution (groups, users) |
| `*.sharepoint.com:443` | SharePoint data access |

### Check Existing EAIs

```sql
SHOW GRANTS TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
-- Look for USAGE on INTEGRATION objects
```

### Create Network Rule + EAI (if needed)

```sql
USE ROLE SECURITYADMIN;

CREATE OR REPLACE NETWORK RULE SHAREPOINT_OPENFLOW_NETWORK_RULE
  TYPE = HOST_PORT
  MODE = EGRESS
  VALUE_LIST = (
    'login.microsoftonline.com:443',
    'login.microsoft.com:443',
    'graph.microsoft.com:443',
    '*.sharepoint.com:443'
  );

CREATE OR REPLACE EXTERNAL ACCESS INTEGRATION SHAREPOINT_OPENFLOW_EAI
  ALLOWED_NETWORK_RULES = (SHAREPOINT_OPENFLOW_NETWORK_RULE)
  ENABLED = TRUE
  COMMENT = 'EAI for Openflow SharePoint connector with Cortex Search and ACLs';

-- Grant to runtime role
GRANT USAGE ON INTEGRATION SHAREPOINT_OPENFLOW_EAI TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
```

### Attach EAI to Runtime

```sql
-- Include any previously attached EAIs in this list (replaces full set)
ALTER OPENFLOW RUNTIME OPENFLOW.OPENFLOW.SPCS1RUNTIME1
  SET EXTERNAL_ACCESS_INTEGRATIONS = (SHAREPOINT_OPENFLOW_EAI, AZURE_SQL_CDC_EAI);
```

> **Note:** If the above SQL fails, attach via the Openflow Control Plane UI:  
> Runtime → "..." menu → "External access integrations" → Select → Save

---

## Step 3: Deploy the Connector

### Option A: Via Openflow UI

1. Navigate to the **Openflow overview page**
2. Click **View more connectors** in the Featured connectors section
3. Find **Microsoft SharePoint (Cortex Search, document ACLs)**
4. Click **Add to runtime** → Select `spcs1runtime1` → Click **Add**
5. Authenticate when prompted

### Option B: Via nipyapi CLI

```bash
# List available SharePoint flows in the Snowflake connector registry
nipyapi --profile tracker_default_spcs1runtime1 ci deploy_flow \
  --bucket "Snowflake Openflow Connector Registry" \
  --flow "unstructured-sharepoint-cdc"
```

After deployment, note the **Process Group ID** returned.

---

## Step 4: Configure Parameters

Right-click the deployed process group → **Parameters**, or use nipyapi:

```bash
# Inspect current parameters
nipyapi --profile tracker_default_spcs1runtime1 ci get_params \
  --process_group_id "<pg-id>"
```

### Set all required parameters:

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci configure_inherited_params \
  --process_group_id "<pg-id>" \
  --parameters '{
    "Sharepoint Site URL": "https://contoso.sharepoint.com/sites/Engineering",
    "Sharepoint Tenant ID": "<your-tenant-id>",
    "Sharepoint Client ID": "<your-client-id>",
    "Sharepoint Client Secret": "<your-client-secret>",
    "Sharepoint Application Certificate": "<paste-full-cert-pem>",
    "Sharepoint Application Private Key": "<paste-full-private-key-pem>",
    "Sharepoint Site Domain": "contoso.sharepoint.com",
    "Destination Database": "SHAREPOINT_DEMO",
    "Destination Schema": "CORTEX",
    "Snowflake Role": "OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1",
    "Snowflake Warehouse": "SI_DEMO",
    "Snowflake Authentication Strategy": "SNOWFLAKE_MANAGED_TOKEN",
    "Snowflake Username": "",
    "Snowflake Account Identifier": "",
    "File Extensions To Ingest": "pdf,docx,pptx,xlsx",
    "Sharepoint Document Library Name": "Shared Documents",
    "Sharepoint Source Folder": "/",
    "Sharepoint Site Groups Enabled": "true",
    "OCR Mode": "LAYOUT",
    "Snowflake Cortex Search Service User Role": "CORTEX_SEARCH_READER"
  }'
```

> **Important notes:**
> - `Snowflake Authentication Strategy` = `SNOWFLAKE_MANAGED_TOKEN` for Snowflake (SPCS) Deployments
> - `Snowflake Username` and `Snowflake Account Identifier` must be **blank** (empty string) for SPCS
> - Certificate/Private Key must include full PEM headers (`-----BEGIN CERTIFICATE-----` etc.)
> - Private key must be **unencrypted**
> - `OCR Mode`: `LAYOUT` preserves table structures as Markdown; `OCR` extracts raw text only
> - `Snowflake Cortex Search Service User Role`: Role granted USAGE on the Cortex Search service

---

## Step 5: Verify Configuration

### Verify Controllers (before enabling)

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci verify_config \
  --process_group_id "<pg-id>" \
  --verify_processors=false
```

> **Expected:** `StandardPrivateKeyService` may show INVALID — this is expected on SPCS (it's for KEY_PAIR auth which is unused). Ignore this warning.

### Enable Controller Services

Right-click on the canvas → **Enable all Controller Services**

Or via nipyapi:
```bash
nipyapi --profile tracker_default_spcs1runtime1 ci enable_controllers \
  --process_group_id "<pg-id>"
```

### Verify Processors (after enabling controllers)

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci verify_config \
  --process_group_id "<pg-id>" \
  --verify_controllers=false
```

---

## Step 6: Start the Flow

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci start_flow \
  --process_group_id "<pg-id>"
```

### Validate

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci get_status \
  --process_group_id "<pg-id>"
```

**Expected:**
- `running_processors` > 0
- `invalid_processors` = 0
- `bulletin_errors` = 0

---

## Step 7: Verify Data Landing

Wait a few minutes for initial ingestion, then check:

```sql
-- Check document chunks (connector creates DOCS_CHUNKS in the destination schema)
SELECT COUNT(*) FROM SHAREPOINT_DEMO.CORTEX.DOCS_CHUNKS;

-- Preview data with ACL columns
SELECT 
  METADATA:fullName::STRING AS file_name,
  METADATA:webUrl::STRING AS web_url,
  LEFT(CHUNK, 200) AS chunk_preview,
  USER_IDS,
  USER_EMAILS
FROM SHAREPOINT_DEMO.CORTEX.DOCS_CHUNKS
LIMIT 10;
```

### Verify Cortex Search Service exists

The connector creates the search service in a schema named `cortex` (lowercase):

```sql
SHOW CORTEX SEARCH SERVICES IN SCHEMA SHAREPOINT_DEMO.cortex;
```

---

## Step 8: Query with ACL Filtering

### SQL Query (with user email filter)

```sql
SELECT PARSE_JSON(
  SNOWFLAKE.CORTEX.SEARCH_PREVIEW(
    'SHAREPOINT_DEMO.cortex.search_service',
    '{
      "query": "What is the vacation policy?",
      "columns": ["chunk", "web_url", "full_name", "user_emails"],
      "filter": {"@contains": {"user_emails": "jane.doe@contoso.com"}},
      "limit": 5
    }'
  )
)['results'] AS results;
```

> **Note:** The connector creates the Cortex Search service as `search_service` in a schema named `cortex` (case-sensitive, lowercase). The fully-qualified reference is `SHAREPOINT_DEMO.cortex.search_service`.

### Python Query (with ACL filter)

```python
from snowflake.core import Root
from snowflake.snowpark import Session

session = Session.builder.configs({
    "connection_name": "default"
}).create()

root = Root(session)

search_service = (root
    .databases["SHAREPOINT_DEMO"]
    .schemas["cortex"]
    .cortex_search_services["search_service"]
)

results = search_service.search(
    query="What is the vacation policy?",
    columns=["chunk", "web_url", "full_name", "user_emails"],
    filter={"@contains": {"user_emails": "jane.doe@contoso.com"}},
    limit=5
)

print(results.to_json())
```

### Query without ACL filter (returns all documents)

```sql
SELECT PARSE_JSON(
  SNOWFLAKE.CORTEX.SEARCH_PREVIEW(
    'SHAREPOINT_DEMO.cortex.search_service',
    '{
      "query": "quarterly financial report",
      "columns": ["chunk", "web_url", "full_name"],
      "limit": 5
    }'
  )
)['results'] AS results;
```

---

## Step 9: (Optional) Grant Access to Other Roles

```sql
-- Create a read-only role for Cortex Search consumers
USE ROLE SECURITYADMIN;
CREATE ROLE IF NOT EXISTS CORTEX_SEARCH_READER;

-- Grant necessary permissions
GRANT USAGE ON DATABASE SHAREPOINT_DEMO TO ROLE CORTEX_SEARCH_READER;
GRANT USAGE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE CORTEX_SEARCH_READER;
GRANT USAGE ON CORTEX SEARCH SERVICE SHAREPOINT_DEMO.cortex.search_service TO ROLE CORTEX_SEARCH_READER;

-- Assign to users
GRANT ROLE CORTEX_SEARCH_READER TO USER <username>;
```

---

## Troubleshooting

| Issue | Cause | Resolution |
|-------|-------|------------|
| `UnknownHostException` in bulletins | Missing EAI domain | Add missing domain to network rule |
| `Authentication failed` | Invalid client secret or certificate | Re-verify credentials in Entra |
| `StandardPrivateKeyService INVALID` | Expected on SPCS | Ignore — not used with session token auth |
| No files syncing | Wrong folder path or library name | Verify `Source Folder` and `Document Library Name` |
| No ACL data (user_ids/user_emails empty) | Certificate not configured or permissions missing | Verify cert + `GroupMember.Read.All` + `User.ReadBasic.All` |
| Cortex Search service not created | Flow hasn't finished initial processing | Wait for first documents to be parsed and chunked |

### Check Bulletins for Errors

```bash
nipyapi --profile tracker_default_spcs1runtime1 bulletins get_bulletin_board
```

### Check Openflow Event Logs

```sql
SELECT *
FROM OPENFLOW.OPENFLOW.EVENTS
WHERE TIMESTAMP > DATEADD('hour', -1, CURRENT_TIMESTAMP())
ORDER BY TIMESTAMP DESC
LIMIT 50;
```

---

## Key Facts

| Item | Value |
|------|-------|
| Connector Flow | `unstructured-sharepoint-cdc` |
| Destination Table | `SHAREPOINT_DEMO.CORTEX.DOCS_CHUNKS` (chunks + ACLs) |
| Cortex Search Service | `SHAREPOINT_DEMO.cortex.search_service` |
| ACL Columns | `user_ids` (ARRAY), `user_emails` (ARRAY) |
| Filter Syntax | `{"@contains": {"user_emails": "user@domain.com"}}` |
| Auth Strategy (SPCS) | `SNOWFLAKE_MANAGED_TOKEN` |
| Runtime Role | `OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1` |
| Runtime URL | `https://of--sfsenorthamerica-demo-tracker.snowflakecomputing.app/spcs1runtime1/nifi-api` |

---

## References

- [Snowflake Docs: SharePoint Connector Setup](https://docs.snowflake.com/en/user-guide/data-integration/openflow/connectors/sharepoint/setup)
- [Snowflake Docs: Query Cortex Search Service](https://docs.snowflake.com/en/user-guide/snowflake-cortex/cortex-search/query-cortex-search-service)
- [Snowflake Docs: Openflow SPCS Domain Allow List](https://docs.snowflake.com/en/user-guide/data-integration/openflow/setup-openflow-spcs-sf-allow-list)
- [Microsoft: Sites.Selected Permission](https://learn.microsoft.com/en-us/graph/permissions-reference#sitesselected)
