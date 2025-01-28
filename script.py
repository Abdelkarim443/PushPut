import asyncio
import csv

from msgraph.core import GraphClient
from azure.identity import ClientSecretCredential

# Azure AD application credentials
client_id = "YOUR_CLIENT_ID"
client_secret = "YOUR_CLIENT_SECRET"
tenant_id = "YOUR_TENANT_ID"

# CSV file paths
input_csv = "mailboxes.csv"
output_csv = "audit_settings.csv"

# Initialize Graph client
credential = ClientSecretCredential(tenant_id, client_id, client_secret)
client = GraphClient(credential=credential)


async def get_mailbox_audit_settings(user_principal_name):
    """Retrieves audit settings for a mailbox."""
    try:
        result = await client.get(
            f"/users/{user_principal_name}/mailboxSettings")
        settings = result.json()
        return {
            "Mailbox": user_principal_name,
            "AuditEnabled": settings.get("auditEnabled"),
            "AuditLogAgeLimit": settings.get("auditLogAgeLimit")
        }
    except Exception as e:
        print(f"Error getting settings for {user_principal_name}: {e}")
        return None


async def main():
    """Main function to orchestrate the process."""
    with open(input_csv, "r") as f:
        reader = csv.reader(f)
        next(reader)  # Skip header
        mailboxes = [row for row in reader]

    tasks = [get_mailbox_audit_settings(mailbox) for mailbox in mailboxes]
    results = await asyncio.gather(*tasks)

    with open(output_csv, "w", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["Mailbox", "AuditEnabled", "AuditLogAgeLimit"])
        writer.writeheader()
        # Filter out any None results (errors)
        writer.writerows(result for result in results if result)


if __name__ == "__main__":
    asyncio.run(main())
