# Cloud PC Audit Event Viewer

A PowerShell-based GUI application for viewing and analyzing Windows 365 Cloud PC audit events from Microsoft Graph.

## Features

- **Interactive GUI** - Windows Forms-based interface for easy data exploration
- **Real-time Data Loading** - Fetches audit events directly from Microsoft Graph API
- **Advanced Filtering** - Filter by activity type, result status, and user principal name
- **Export Capabilities** - Export filtered results to CSV format
- **Progress Tracking** - Visual progress bar with percentage indicator during data loading
- **Responsive Design** - Resizable window with anchored controls for optimal viewing

## Prerequisites

### Required Permissions
- **CloudPC.Read.All** - Microsoft Graph API scope for reading Cloud PC audit events

### PowerShell Modules
- **Microsoft.Graph.Beta** - For accessing Cloud PC audit event data
  ```powershell
  Install-Module Microsoft.Graph.Beta -Scope CurrentUser
  ```

### System Requirements
- PowerShell 5.1 or higher
- Windows OS with .NET Framework
- Active Microsoft 365/Azure AD tenant with Windows 365 license

## Installation

1. Clone or download this repository
2. Ensure the Microsoft Graph PowerShell SDK is installed
3. Run the script with appropriate permissions:
   ```powershell
   .\W365_AuditGUI.ps1
   ```

## Usage

### Initial Connection
1. Click **Load Data** button
2. If not already connected, you'll be prompted to authenticate to Microsoft Graph
3. Grant the required **CloudPC.Read.All** permission when prompted
4. Wait for the data to load (progress bar will indicate status)

### Filtering Data

#### By Activity Type
- Use the **Activity Type** dropdown to filter by specific event types
- Options populate dynamically based on loaded data

#### By Activity Result
- Filter by **Success** or **Failure** status
- Default shows all results

#### By User
- Enter a username or email in the **Search (UPN)** field
- Supports partial matching with wildcard behavior
- Press Enter or click **Apply Filter** to execute

### Exporting Data
1. Load and/or filter the data as desired
2. Click **Export CSV** button
3. Choose save location and filename
4. File is exported with timestamp: `CloudPC_AuditEvents_YYYYMMDD_HHMMSS.csv`

### Clearing Filters
- Click **Clear Filter** to reset all filters and show complete dataset

## Data Columns

The application displays the following information for each audit event:

| Column | Description |
|--------|-------------|
| **Activity DateTime** | Timestamp when the activity occurred |
| **Activity Type** | Type of activity performed (e.g., provision, deprovision) |
| **Activity Result** | Success or Failure status |
| **User Id** | Azure AD User ID (GUID) |
| **User Principal Name** | User's email/UPN |

## UI Components

- **Status Label** - Displays current operation status and messages
- **Progress Bar** - Shows loading progress with percentage
- **Load Data Button** - Fetches fresh data from Microsoft Graph
- **Apply Filter Button** - Applies selected filter criteria
- **Clear Filter Button** - Resets all filters
- **Export CSV Button** - Exports current view to CSV file
- **DataGridView** - Main data display with sortable columns and row selection

## Error Handling

The application includes comprehensive error handling for:
- Authentication failures
- Network connectivity issues
- API rate limiting
- Invalid data scenarios
- File export errors

Error messages are displayed via:
- Status label (colored indicators)
- Message box dialogs for critical errors

## Authentication

The script uses **Connect-MgGraph** with device code flow authentication. You'll need:
- Valid Azure AD credentials
- Appropriate role assignment (Global Reader, Cloud PC Administrator, or custom role with CloudPC read permissions)

## Customization

### Modify Time Range
By default, the script fetches all available audit events. To filter by date range, modify the `Get-MgBetaDeviceManagementVirtualEndpointAuditEvent` command:

```powershell
$auditEvents = Get-MgBetaDeviceManagementVirtualEndpointAuditEvent -All -Filter "activityDateTime ge $startDate and activityDateTime le $endDate"
```

### Adjust Window Size
Modify the initial form size in the code:
```powershell
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 600)
```

### Add Additional Columns
To display more audit event properties, expand the `Select-Object` and `Update-DataGrid` functions to include additional fields available from the Graph API.

## Troubleshooting

### "Not connected to Microsoft Graph"
- Ensure Microsoft.Graph.Beta module is installed
- Check internet connectivity
- Verify tenant access and permissions

### "No audit events found"
- Confirm Windows 365 Cloud PCs are provisioned in the tenant
- Verify audit logging is enabled
- Check date range if filtering is applied

### Performance Issues
- Large datasets (>10,000 records) may take longer to load
- Consider adding date filters for specific time ranges
- Use the filtering features to reduce displayed data

## License

This script is provided as-is for use within organizations using Windows 365 Cloud PC services.

### Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The author(s) assume no responsibility for any damage, data loss, or misuse resulting from the use of this script. Users are responsible for ensuring proper testing in non-production environments and compliance with their organization's security policies and applicable regulations.

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## Author

Created for Windows 365 Cloud PC audit event analysis and monitoring.

## Version History

- **1.0** - Initial release with core functionality
  - Data loading from Microsoft Graph
  - Filtering by activity type, result, and user
  - CSV export capability
  - Progress tracking
  - Responsive UI design

## Support

For issues related to:
- **Microsoft Graph API** - Refer to [Microsoft Graph Documentation](https://learn.microsoft.com/en-us/graph/)
- **Windows 365** - Visit [Windows 365 Documentation](https://learn.microsoft.com/en-us/windows-365/)
- **PowerShell SDK** - Check [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
