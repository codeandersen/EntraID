# Login if needed
Connect-AzAccount

# Get ARM access token
$token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token

# Build the URL
$uri = "https://management.azure.com/subscriptions/5e480aad-e2f1-4f9c-b6b2-ef906d0c8844/resourceGroups/rg-email-communication-services-stark-suomi-fi-prod/providers/Microsoft.Communication/emailServices/acs-mail-stark-suomi-fi-prod/domains/email.stark-suomi.fi/senderUsernames/uutiskirja?api-version=2025-05-01"

# Send the DELETE request
$response = Invoke-RestMethod -Method Delete -Uri $uri -Headers @{ Authorization = "Bearer $token" }

$response