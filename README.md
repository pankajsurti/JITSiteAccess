# JITSiteAccess
The repo is a home for the Just In Time (JIT) SharePoint Site Access solution.

## Common Customer Scenarios

- Site Collection Admins may ask for help from colleagues, help desk and other Site Collection Admins.
- In SharePoint online, there is no Web App Policy (like there is on-premises) to allow access to all Sites.
- SharePoint Administrators can add themselves to the Site Collection Administrator for an individual site through SharePoint Admin Center.
- SharePoint Administrators must manually remove themselves from each site they have been added to if required.

## Solution

An Automation to add and remove Privledged Admin users for defined period of time. Just In Time (JIT) Site Admins.

## Step By Step Implementation

### Step 1: Create a Site Collection and a following list

| InternalName | Display Name |Field Type| ShowInForm
| ----------- | ----------- |---|----------|
| Title | Site Name |SingleText| Yes |
| reqJustification |Note| Justification | Yes |
| reqStatus | Request Status |Choice| No |
| reqSysStatus | System Status |Choice| No |
| reqActivateTime | Activate Time |DateTime| No |
| reqExpiryTime | Expiry Time |DateTime| No |
| reqExpireTimeMin | Expire Time in Minutes |Number| No |
| reqApprovers | Site Approvers |Multi Person| No |

### Step 2: Create Azure Function

Follow [How to setup certificate in MAG Azure Function App for PnP.PowerShell?](https://pankajsurti.wordpress.com/2021/06/11/how-to-setup-certificate-in-mag-azure-function-app-for-pnp-powershell/)
