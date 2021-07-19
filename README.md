# Just In Time (JIT) Site Access
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
### Architecture

![image](https://user-images.githubusercontent.com/108044/126233431-6baff594-f229-4611-8e69-f25b5e570b53.png)

For the JIT solution the architecture uses two SaaS clouds. ​

1. For the Office 365 side, the SharePoint plays an important part to collect the Site Admin request in secured way. Only the privileged admin users are allowed to access the site and tool. 

2. For the Azure side, there are two important services are playing the role of the JIT Engine. The Azure Key Vault is used to securely store the PFX certificate used in communicating to the SharePoint Tenant using the PnP PowerShell. The Azure Function is a compute service which runs as a timer service to serve the requests in the SharePoint List.

3. The Power Automate plays a role of the approval flows and notification to the requestor.

4. For the simplicity we have chosen to use the Out Of the Box SharePoint Entry form but this solution can be extended to use PowerApps for the entry forms.

![image](https://user-images.githubusercontent.com/108044/126233589-57a211c8-f095-4e5b-853a-5c10840a687d.png)

There are two fields in the SharePoint List which makes it possible to do the communication between Power Automate and the Azure Function.​

1. System Status

2. Request Status


When user enters a new request the Request Status field defaults to the "Pending" state and the System Status defaults to "NEWREQUEST"​

The JIT Engine runs a defined internal picks up the "Pending" request. The request is populated with the list of Site Collection Admins. This SCA list is used in the approval flow later step.

At this time, the Request Status is changed to Approved and System Status is changed to "TRIGGER-ADMIN-UPDATE". ​

This change will trigger the Power Automate flow on the item modified event. The flow will send an approval emails to the SCAs​

The SCA will approve or reject for the approval message.​

If approved, the JIT engine will pick up the reqest when it runs. It will change add the User to the requested site as a site collection admin. ​

It changes the REquest Status to Active and System Status to TRIGGER-USER-ACTIVE. This item modified event will fire another Flow to send an email to the requestor that site is accesible as an SCA.

JIT Engine looks for the Active status and Expiry time. If expired, the JIT Engine changes the Request Status to Completed and System Status to TRIGGER-USER-REMOVED

In case of any error, Request Status to Error and System Status to TRIGGER-ERROR

In case of Rejection, Request Status to Rejected and System Status to TRIGGER-ADMIN-REJECTED​



