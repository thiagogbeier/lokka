# Lokka Usage Examples

This document provides examples of how to use Lokka's tools with natural language queries and direct API calls.

## Getting Recent Groups

### Use Case: Query Groups Created in the Past Week

The `get-recent-groups` tool makes it easy to find groups created within a specific time period.

#### Natural Language Query Examples

When using Lokka with an AI assistant (like Claude or ChatGPT), you can use natural language:

- "Show me all groups created in the past week"
- "List groups created in the last 7 days"
- "What groups were created in the past 3 days?"
- "Find recently created groups from the last 14 days"

#### Direct Tool Usage

If calling the tool directly through MCP, use these parameters:

**Example 1: Get groups from the past week (default)**
```json
{
  "tool": "get-recent-groups"
}
```

**Example 2: Get groups from the past 3 days**
```json
{
  "tool": "get-recent-groups",
  "daysAgo": 3
}
```

**Example 3: Get groups from past 30 days with all fields**
```json
{
  "tool": "get-recent-groups",
  "daysAgo": 30,
  "includeAllFields": true
}
```

**Example 4: Get first page only from past 14 days**
```json
{
  "tool": "get-recent-groups",
  "daysAgo": 14,
  "fetchAll": false
}
```

### Expected Response

The tool returns:
- A summary including the query filter, threshold date, and total count
- Full JSON response with group details including:
  - `id`: Group's unique identifier
  - `displayName`: Group's display name
  - `createdDateTime`: When the group was created
  - `description`: Group description (if any)
  - `groupTypes`: Array indicating group types (e.g., "Unified" for Microsoft 365 groups)
  - `mail`: Group's email address (if any)

### Required Permissions

To use this tool, the authenticated user or application needs one of these Microsoft Graph permissions:
- `Group.Read.All` (read-only access)
- `Group.ReadWrite.All` (read-write access)
- `Directory.Read.All` (read all directory data)

### Implementation Details

Under the hood, the tool:
1. Calculates the threshold date (current date minus `daysAgo`)
2. Constructs a Microsoft Graph filter: `createdDateTime ge [threshold-date]`
3. Queries the `/groups` endpoint with the filter
4. Sorts results by `createdDateTime` in descending order (newest first)
5. Uses `ConsistencyLevel: eventual` header for advanced query support
6. Automatically handles pagination if `fetchAll` is true

### Alternative: Using Lokka-Microsoft Tool Directly

You can also use the general-purpose `Lokka-Microsoft` tool to achieve the same result:

```json
{
  "tool": "Lokka-Microsoft",
  "apiType": "graph",
  "path": "/groups",
  "method": "get",
  "queryParams": {
    "$filter": "createdDateTime ge 2025-10-24T00:00:00Z",
    "$orderby": "createdDateTime desc",
    "$select": "id,displayName,createdDateTime,description,groupTypes,mail"
  },
  "consistencyLevel": "eventual",
  "fetchAll": true
}
```

Note: You would need to manually calculate the date for the `$filter` parameter.

## Tips

1. **Performance**: For large tenants with many groups, consider:
   - Using `fetchAll: false` to get just the first page
   - Reducing the time window (e.g., 3 days instead of 30)

2. **Troubleshooting**: If you get permission errors:
   - Verify the app/user has the required Graph permissions
   - Check if admin consent has been granted for app permissions
   - Use `get-auth-status` tool to check current authentication state

3. **Date Filtering**: The tool uses ISO 8601 format for dates and the `ge` (greater than or equal) operator, which is efficient for querying recent items.
