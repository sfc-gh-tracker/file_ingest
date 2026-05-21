# Workshop: Setting Up the Openflow SharePoint Connector with Cortex Search and Document ACLs

> **Goal:** Walk through every step of configuring the SharePoint ACL connector on an SPCS-managed Openflow runtime—explaining what each configuration does and why it matters.

| | |
|---|---|
| **Deployment type** | SPCS (Snowflake-managed runtime) |
| **Connector variant** | `unstructured-sharepoint-cdc` (Cortex Search + Document ACLs) |
| **Runtime** | `spcs1runtime1` |

---

## How It Works

```
┌──────────────────┐       ┌─────────────────────────┐       ┌──────────────────────────┐
│  SharePoint Site │──────▶│  Openflow SPCS Runtime  │──────▶│  Snowflake Cortex Search │
│  (Documents +    │       │  (NiFi-based connector) │       │  (with ACL filtering)    │
│   Permissions)   │       └─────────────────────────┘       └──────────────────────────┘
└──────────────────┘
```

The connector:
1. Connects to your SharePoint site via Microsoft Graph API
2. Pulls documents from the document library you specify
3. Parses each document with `AI_PARSE_DOCUMENT` (Snowflake's built-in document OCR)
4. Chunks the parsed text and writes it to a `DOCS_CHUNKS` table
5. Resolves document-level permissions (who can access each file) and stores them as `user_ids` and `user_emails` arrays alongside each chunk
6. Creates and maintains a Cortex Search Service that indexes the chunks
7. Runs continuously via CDC — new/updated/deleted files are synced automatically

When you query the Cortex Search service, you pass a `user_emails` filter so results are scoped to only documents that user can access in SharePoint.

---

## Understanding How ACLs Work (The Trust Model)

This is important to understand before configuring anything. The ACL mechanism is **application-level filtering** — it is NOT SAML integration, Snowflake role mapping, or automatic enforcement.

### What Happens at Ingestion Time

```
SharePoint                              Snowflake (DOCS_CHUNKS table)
─────────                              ─────────────────────────────
Document "policy.pdf"                   Row for "policy.pdf"
  Permissions:                            USER_EMAILS: ["jane@contoso.com",
    - Jane (direct access)                              "bob@contoso.com",
    - "Engineering" group                               "alice@contoso.com"]
        - Bob                             USER_IDS:    ["guid-jane", "guid-bob",
        - Alice                                         "guid-alice"]
```

The connector:
1. Reads each document's SharePoint permissions (which users and groups are assigned)
2. Calls Microsoft Graph API to expand groups into individual members (`GroupMember.Read.All`)
3. Calls Microsoft Graph API to resolve user GUIDs into email addresses (`User.ReadBasic.All`)
4. Stores the **flattened** list of user emails and IDs as array columns alongside each document chunk

### What Happens at Query Time

```
User logs into your app (via Entra SSO, SAML, whatever you use)
        │
        ▼
Your app knows the user's email (e.g., from the session/token)
        │
        ▼
Your app passes that email as a filter to Cortex Search
        │
        ▼
Cortex Search returns only chunks where user_emails array contains that email
```

The `@contains` filter checks: "is this email present in the `user_emails` array for this chunk?" If yes, the chunk is returned. If no, it's excluded.

### What This Is NOT

| Common Misconception | Reality |
|---|---|
| Snowflake roles mapped to Entra groups | No. The ACL is just data in an array column — no Snowflake RBAC integration. |
| SAML/SSO enforcement at the Snowflake layer | No. There is no identity federation between Entra and Snowflake for search results. |
| Automatically enforced by Snowflake | No. Anyone with USAGE on the search service can query without a filter and see everything. Your application must pass the filter. |
| Real-time permission sync | No. Permissions sync via CDC — there is a small delay (seconds to minutes) between a SharePoint permission change and the ACL array being updated in Snowflake. |

### The Security Boundary

**Your application is the enforcement layer**, not Snowflake. Snowflake faithfully stores the ACL data and Cortex Search efficiently applies the filter, but it is the calling code that decides which email to filter on.

This means:
- An admin querying directly in a SQL worksheet (without a filter) sees all documents
- Your app must authenticate the user independently and inject the correct email into the filter
- If you need Snowflake-level enforcement, you would wrap the search service call in a stored procedure or UDF that injects the filter based on `CURRENT_USER()` email mapping

---

## Part 1: Microsoft Entra ID Setup

Before touching Snowflake, you need an App Registration in Microsoft Entra ID (formerly Azure AD). This is how the connector authenticates to SharePoint and resolves permissions.

### 1.1 Create (or reuse) an App Registration

In the Azure Portal → **Microsoft Entra ID** → **App registrations** → **New registration**.

Record these values — you'll need them later:

| Value | Where to find it | What it's for |
|-------|-------------------|---------------|
| **Tenant ID** | Entra Admin Center → Overview | Identifies your Microsoft 365 tenant |
| **Client ID** | App Registration → Overview → Application (client) ID | Identifies this specific app |
| **Client Secret** | App Registration → Certificates & secrets → New client secret | Password-based auth for API calls |

### 1.2 Configure API Permissions

Navigate to **App Registration → API permissions → Add a permission**.

The ACL connector requires these **Application** permissions (not Delegated):

| Permission | API | Why the connector needs it |
|------------|-----|---------------------------|
| `Sites.Selected` | Microsoft Graph | Scopes access to only the sites you explicitly authorize. This is the least-privilege approach — the connector cannot access any site you haven't granted. |
| `GroupMember.Read.All` | Microsoft Graph | When a SharePoint document is shared with a group, the connector needs to resolve group membership to individual users. Without this, group-based permissions won't flow through to the ACL columns. |
| `User.ReadBasic.All` | Microsoft Graph | Resolves Microsoft 365 user object IDs to email addresses. The `user_emails` column in your search results comes from this permission. |
| `Sites.Selected` | SharePoint | Required in addition to the Graph-level Sites.Selected. Microsoft requires it in both places for full site access. |

After adding all four, click **Grant admin consent** (requires Global Admin or Privileged Role Administrator).

### 1.3 Grant the App Access to Your Specific Site

`Sites.Selected` alone doesn't grant access — it just enables the pattern. You must explicitly grant the app a role on each site.

**Why `fullcontrol`?** The connector needs to detect permission changes on folders during CDC. If a folder's permissions change and the connector only has `read`, it cannot detect the change and may enter an irreparable state requiring full re-ingestion. `fullcontrol` prevents this.

**Option A: PowerShell (PnP module)**

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<sitename>" -Interactive

Grant-PnPAzureADAppSitePermission `
  -AppId "<client-id>" `
  -DisplayName "Openflow SharePoint Connector" `
  -Permissions FullControl `
  -Site "https://<tenant>.sharepoint.com/sites/<sitename>"
```

**Option B: Microsoft Graph API (curl)**

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

### 1.4 Generate a Certificate and Private Key

The ACL connector uses certificate-based auth specifically for resolving group memberships and user lookups. This is separate from the client secret (which handles the main SharePoint API calls).

```bash
openssl req -x509 -nodes -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 \
  -subj "/CN=OpenflowSharePointConnector"
```

| File | What it is | Where it goes |
|------|------------|---------------|
| `cert.pem` | Public certificate | Upload to App Registration → Certificates & secrets → Certificates |
| `key.pem` | Private key (unencrypted) | Paste into the `Sharepoint Application Private Key` parameter in Openflow |

> **The `-nodes` flag** means "no DES encryption" on the private key. The connector requires an unencrypted key.

---

## Part 2: Snowflake Account Setup

### 2.1 Create the Destination Database and Schema

The connector writes parsed document chunks and metadata here. These must exist before you start the flow.

```sql
USE ROLE ACCOUNTADMIN;

CREATE DATABASE IF NOT EXISTS SHAREPOINT_DEMO;
CREATE SCHEMA IF NOT EXISTS SHAREPOINT_DEMO.CORTEX;
```

### 2.2 Grant Privileges to the Runtime Role

The Openflow runtime operates under a dedicated Snowflake role. That role needs permissions to create the objects the connector will produce.

```sql
-- The runtime role for your SPCS runtime (find it in Openflow UI → Runtime → View Details)
-- For this workshop: OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1

GRANT USAGE ON DATABASE SHAREPOINT_DEMO TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT USAGE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE TABLE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE DYNAMIC TABLE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE STAGE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE SEQUENCE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT CREATE CORTEX SEARCH SERVICE ON SCHEMA SHAREPOINT_DEMO.CORTEX TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
GRANT USAGE, OPERATE ON WAREHOUSE SI_DEMO TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
```

**Why each grant matters:**

| Grant | What the connector creates with it |
|-------|-----------------------------------|
| `CREATE TABLE` | `DOCS_CHUNKS` table (document chunks + ACLs), file hash tracking table |
| `CREATE DYNAMIC TABLE` | Internal CDC state tables |
| `CREATE STAGE` | Internal stage for raw file storage |
| `CREATE SEQUENCE` | Sequence generators for chunk IDs |
| `CREATE CORTEX SEARCH SERVICE` | The search service that indexes your documents |
| `USAGE, OPERATE ON WAREHOUSE` | Runs the SQL queries that write data and refresh the search service |

---

## Part 3: External Access Integration (EAI)

### Why this is needed

SPCS runtimes run inside Snowflake's network. By default, they **cannot** reach any external endpoint. The EAI is Snowflake's allowlist mechanism — you explicitly declare which external hosts the runtime is permitted to contact.

### 3.1 Understand the Required Domains

| Domain | Port | Why the connector contacts it |
|--------|------|-------------------------------|
| `login.microsoftonline.com` | 443 | OAuth2 token endpoint — authenticates the app to get access tokens |
| `login.microsoft.com` | 443 | Fallback OAuth endpoint (some tenants route here) |
| `graph.microsoft.com` | 443 | Microsoft Graph API — resolves group memberships and user emails for ACLs |
| `*.sharepoint.com` | 443 | The actual SharePoint site data (file content, metadata, permissions). The wildcard covers any tenant's subdomain. |

### 3.2 Check What Already Exists

```sql
-- See what integrations the runtime role already has access to
SHOW GRANTS TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
```

If you already have an EAI covering these domains (perhaps from a previous connector), you can skip creation.

### 3.3 Create the Network Rule

A Network Rule defines the allowlist of host:port combinations. The EAI then references this rule.

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
```

**Key details:**
- `TYPE = HOST_PORT` — controls which DNS names + ports can be reached
- `MODE = EGRESS` — this is outbound traffic (runtime → internet)
- Every entry **must** include the port (`:443`). Without a port, the rule won't work.
- `*.sharepoint.com` — wildcard matches one subdomain level (e.g., `contoso.sharepoint.com`)

### 3.4 Create the External Access Integration

```sql
CREATE OR REPLACE EXTERNAL ACCESS INTEGRATION SHAREPOINT_OPENFLOW_EAI
  ALLOWED_NETWORK_RULES = (SHAREPOINT_OPENFLOW_NETWORK_RULE)
  ENABLED = TRUE
  COMMENT = 'EAI for Openflow SharePoint connector with Cortex Search and ACLs';
```

### 3.5 Grant USAGE to the Runtime Role

```sql
GRANT USAGE ON INTEGRATION SHAREPOINT_OPENFLOW_EAI TO ROLE OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1;
```

### 3.6 Attach the EAI to the Runtime

```sql
-- This REPLACES the full list — include any previously attached EAIs
ALTER OPENFLOW RUNTIME OPENFLOW.OPENFLOW.SPCS1RUNTIME1
  SET EXTERNAL_ACCESS_INTEGRATIONS = (SHAREPOINT_OPENFLOW_EAI, AZURE_SQL_CDC_EAI);
```

> **If this SQL fails** (e.g., on older accounts), do it through the Openflow Control Plane UI:  
> Runtime → `...` menu → "External access integrations" → Select from dropdown → Save

No restart required — changes apply immediately.

---

## Part 4: Deploy the Connector

### 4.1 Install via Openflow UI

1. Open the **Openflow overview page** in Snowsight
2. Click **View more connectors**
3. Find **"Microsoft SharePoint (Cortex Search, document ACLs)"**
4. Click **Add to runtime** → select your runtime → **Add**
5. Authenticate when prompted (this grants the connector access to your Snowflake session)

The connector appears as a process group on the Openflow canvas.

### 4.2 Alternative: Deploy via CLI

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci deploy_flow \
  --bucket "Snowflake Openflow Connector Registry" \
  --flow "unstructured-sharepoint-cdc"
```

Note the **Process Group ID** returned — you'll use it in all subsequent commands.

---

## Part 5: Configure Parameters (The Important Part)

This is where most questions come up. Each parameter has a specific purpose. Right-click the process group → **Parameters** in the UI, or use the CLI.

### 5.1 Parameter Reference

#### Source Parameters (SharePoint connection)

| Parameter | Example Value | What it does |
|-----------|---------------|--------------|
| `Sharepoint Site URL` | `https://contoso.sharepoint.com/sites/Engineering` | The full URL of the SharePoint site to ingest from. Must include `/sites/<name>`. |
| `Sharepoint Tenant ID` | `72f988bf-86f1-41af-91ab-2d7cd011db47` | Your Microsoft 365 tenant ID. Used in OAuth token requests. |
| `Sharepoint Client ID` | `a1b2c3d4-e5f6-...` | The Application (client) ID from your App Registration. Identifies the app to Microsoft. |
| `Sharepoint Client Secret` | `(sensitive)` | The client secret value (not the secret ID). Used for the main API authentication flow. |
| `Sharepoint Application Certificate` | `-----BEGIN CERTIFICATE-----\nMIIC...` | The full PEM certificate. Used specifically for certificate-based auth to resolve ACLs. |
| `Sharepoint Application Private Key` | `-----BEGIN PRIVATE KEY-----\nMIIE...` | The corresponding private key (unencrypted). Paired with the certificate above. |
| `Sharepoint Site Domain` | `contoso.sharepoint.com` | Just the domain portion of your site URL. Used for constructing Graph API calls for ACL resolution. |

#### Ingestion Parameters (what to pull)

| Parameter | Example Value | What it does |
|-----------|---------------|--------------|
| `Sharepoint Source Folder` | `/` | Folder path relative to the document library root. `/` means ingest everything. `/Engineering/Specs` would only ingest that subfolder (and its children). |
| `File Extensions To Ingest` | `pdf,docx,pptx,xlsx` | Comma-separated list. Only files with these extensions are ingested. Empty string = all supported files. The connector attempts PDF conversion first. |
| `Sharepoint Document Library Name` | `Shared Documents` | The name of the SharePoint document library to read from. Most sites use the default "Shared Documents". |
| `Sharepoint Site Groups Enabled` | `true` | When `true`, the connector resolves SharePoint site group memberships into individual user IDs/emails in the ACL columns. Set to `false` if you don't use site groups for permissions. |
| `OCR Mode` | `LAYOUT` | Controls how `AI_PARSE_DOCUMENT` processes files. `LAYOUT` preserves table structures as Markdown (better for structured docs). `OCR` extracts raw text only (faster, but loses formatting). |
| `Snowflake File Hash Table Name` | (leave default) | Internal tracking table for change detection. Don't modify unless you know what you're doing. |

#### Destination Parameters (where to write in Snowflake)

| Parameter | Example Value | What it does |
|-----------|---------------|--------------|
| `Destination Database` | `SHAREPOINT_DEMO` | The database you created in Part 2. Must already exist. **Case-sensitive** — use uppercase for unquoted identifiers. |
| `Destination Schema` | `CORTEX` | The schema you created in Part 2. Must already exist. **Case-sensitive** — use uppercase for unquoted identifiers. |
| `Snowflake Role` | `OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1` | The Snowflake role the connector operates as. On SPCS, this is the runtime role shown in Openflow UI → Runtime → View Details. |
| `Snowflake Warehouse` | `SI_DEMO` | Warehouse used for SQL queries (creating tables, writing chunks, refreshing the search service). |
| `Snowflake Authentication Strategy` | `SNOWFLAKE_MANAGED_TOKEN` | How the connector authenticates to Snowflake. On SPCS deployments, **always** use `SNOWFLAKE_MANAGED_TOKEN` — the token is handled automatically by the platform. |
| `Snowflake Username` | *(blank)* | Leave empty for SPCS. Only used with `KEY_PAIR` auth (BYOC deployments). |
| `Snowflake Account Identifier` | *(blank)* | Leave empty for SPCS. Only used with `KEY_PAIR` auth. |
| `Snowflake Cortex Search Service User Role` | `CORTEX_SEARCH_READER` | The role that will be granted USAGE on the Cortex Search service the connector creates. This is who can query the service. |

### 5.2 Set Parameters via CLI

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci configure_inherited_params \
  --process_group_id "<pg-id>" \
  --parameters '{
    "Sharepoint Site URL": "https://contoso.sharepoint.com/sites/Engineering",
    "Sharepoint Tenant ID": "<your-tenant-id>",
    "Sharepoint Client ID": "<your-client-id>",
    "Sharepoint Client Secret": "<your-client-secret>",
    "Sharepoint Application Certificate": "<full-cert-pem-content>",
    "Sharepoint Application Private Key": "<full-private-key-pem-content>",
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

---

## Part 6: Verify and Start

### 6.1 Verify Controller Configuration

Controllers are background services (connection pools, auth handlers). Verify them before enabling:

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci verify_config \
  --process_group_id "<pg-id>" \
  --verify_processors=false
```

> **Expected warning:** `StandardPrivateKeyService` shows INVALID. This is normal on SPCS — that controller is for KEY_PAIR auth (used in BYOC). It's unused here. Ignore it.

### 6.2 Enable Controller Services

In the UI: Right-click the canvas → **Enable all Controller Services**

Or via CLI:
```bash
nipyapi --profile tracker_default_spcs1runtime1 ci enable_controllers \
  --process_group_id "<pg-id>"
```

### 6.3 Verify Processors

Processors are the actual data-processing units. They depend on controllers being enabled first:

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci verify_config \
  --process_group_id "<pg-id>" \
  --verify_controllers=false
```

### 6.4 Start the Flow

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci start_flow \
  --process_group_id "<pg-id>"
```

### 6.5 Confirm It's Running

```bash
nipyapi --profile tracker_default_spcs1runtime1 ci get_status \
  --process_group_id "<pg-id>"
```

**What to look for:**
- `running_processors` > 0 — processors are active
- `invalid_processors` = 0 — no configuration issues
- `bulletin_errors` = 0 — no runtime errors

---

## Part 7: Verify Data Is Landing

Give it a few minutes for the initial sync, then check:

### 7.1 Check the Chunks Table

```sql
SELECT COUNT(*) FROM SHAREPOINT_DEMO.CORTEX.DOCS_CHUNKS;
```

### 7.2 Inspect the Data Structure

```sql
SELECT 
  METADATA:fullName::STRING AS file_name,
  METADATA:webUrl::STRING AS web_url,
  LEFT(CHUNK, 200) AS chunk_preview,
  USER_IDS,
  USER_EMAILS
FROM SHAREPOINT_DEMO.CORTEX.DOCS_CHUNKS
LIMIT 10;
```

**What you're seeing:**
- `METADATA:fullName` — the file path in SharePoint (e.g., `HR/Policies/vacation.pdf`)
- `CHUNK` — a text segment from the parsed document
- `USER_IDS` — array of Microsoft 365 user GUIDs who can access this document
- `USER_EMAILS` — array of email addresses (resolved from user IDs + group memberships)

### 7.3 Verify the Cortex Search Service

The connector automatically creates a search service named `search_service` in a schema it creates called `cortex` (lowercase):

```sql
SHOW CORTEX SEARCH SERVICES IN SCHEMA SHAREPOINT_DEMO.cortex;
```

---

## Part 8: Querying with ACL Filtering

This is where the ACL data becomes useful. By passing a filter, you restrict search results to only documents the specified user has access to in SharePoint.

### 8.1 SQL Query with ACL Filter

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

**How the filter works:**
- `@contains` checks if the `user_emails` array contains the specified email
- Only chunks from documents where that email appears in the ACL are returned
- This mirrors the SharePoint permission model — if Jane can't access a file in SharePoint, she won't see it in search results

### 8.2 Python Query

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

### 8.3 Query Without Filter (Admin View)

Omitting the filter returns results from all documents regardless of permissions:

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

### 8.4 Available Columns in Search Results

| Column | Type | Description |
|--------|------|-------------|
| `chunk` | String | The text segment that matched the search query |
| `full_name` | String | Full file path from the site documents root (e.g., `folder_1/file.pdf`) |
| `web_url` | String | Direct URL to open the original file in SharePoint |
| `last_modified_date_time` | String | When the file was last modified |
| `user_ids` | Array | Microsoft 365 user GUIDs with access to this document |
| `user_emails` | Array | Email addresses with access (includes resolved group memberships) |

---

## Part 9: Grant Search Access to Other Roles

The `Snowflake Cortex Search Service User Role` parameter (set in Part 5) controls who can query the search service. If you need additional roles:

```sql
USE ROLE SECURITYADMIN;
CREATE ROLE IF NOT EXISTS CORTEX_SEARCH_READER;

GRANT USAGE ON DATABASE SHAREPOINT_DEMO TO ROLE CORTEX_SEARCH_READER;
GRANT USAGE ON SCHEMA SHAREPOINT_DEMO.cortex TO ROLE CORTEX_SEARCH_READER;
GRANT USAGE ON CORTEX SEARCH SERVICE SHAREPOINT_DEMO.cortex.search_service TO ROLE CORTEX_SEARCH_READER;

GRANT ROLE CORTEX_SEARCH_READER TO USER <username>;
```

---

## Troubleshooting

| Symptom | What's happening | How to fix |
|---------|-----------------|------------|
| `UnknownHostException` in bulletins | The runtime can't resolve a hostname — it's not in the EAI | Check which domain failed, add it to the network rule |
| `Authentication failed` | Client secret expired, or certificate doesn't match what's registered | Regenerate credentials in Entra, update the parameter |
| `StandardPrivateKeyService INVALID` | Normal on SPCS | Ignore — this controller is for KEY_PAIR auth only |
| No files appearing after start | Source folder path doesn't match, or document library name is wrong | Double-check `Sharepoint Source Folder` and `Sharepoint Document Library Name` in parameters |
| `user_ids`/`user_emails` arrays are empty | Certificate/private key not configured, or Graph permissions missing | Verify the cert is uploaded to Entra AND pasted in parameters; verify `GroupMember.Read.All` + `User.ReadBasic.All` have admin consent |
| Cortex Search service doesn't appear | The flow hasn't finished initial document processing yet | Wait longer — first run parses all docs before creating the service |
| `SocketTimeoutException` | DNS resolves but TCP connection fails — port may be missing from network rule | Verify every VALUE_LIST entry includes `:443` |

### Check Bulletins

```bash
nipyapi --profile tracker_default_spcs1runtime1 bulletins get_bulletin_board
```

### Check Event Logs

```sql
SELECT *
FROM OPENFLOW.OPENFLOW.EVENTS
WHERE TIMESTAMP > DATEADD('hour', -1, CURRENT_TIMESTAMP())
ORDER BY TIMESTAMP DESC
LIMIT 50;
```

---

## Quick Reference

| Item | Value |
|------|-------|
| Connector flow name | `unstructured-sharepoint-cdc` |
| Destination table | `SHAREPOINT_DEMO.CORTEX.DOCS_CHUNKS` |
| Cortex Search service | `SHAREPOINT_DEMO.cortex.search_service` |
| ACL columns | `user_ids` (ARRAY), `user_emails` (ARRAY) |
| ACL filter syntax | `{"@contains": {"user_emails": "user@domain.com"}}` |
| Auth strategy (SPCS) | `SNOWFLAKE_MANAGED_TOKEN` |
| Runtime role | `OPENFLOWRUNTIMEROLE_SPCS1_RUNTIME1` |

---

## Further Reading

- [Snowflake Docs: SharePoint Connector Setup](https://docs.snowflake.com/en/user-guide/data-integration/openflow/connectors/sharepoint/setup)
- [Snowflake Docs: About the SharePoint Connector](https://docs.snowflake.com/en/user-guide/data-integration/openflow/connectors/sharepoint/about)
- [Snowflake Docs: Query a Cortex Search Service](https://docs.snowflake.com/en/user-guide/snowflake-cortex/cortex-search/query-cortex-search-service)
- [Snowflake Docs: Openflow SPCS Domain Allow List](https://docs.snowflake.com/en/user-guide/data-integration/openflow/setup-openflow-spcs-sf-allow-list)
- [Microsoft: Sites.Selected Permission](https://learn.microsoft.com/en-us/graph/permissions-reference#sitesselected)
- [Microsoft: Grant App Site Permission (PnP)](https://github.com/pnp/powershell/blob/dev/documentation/Grant-PnPAzureADAppSitePermission.md)
