/*output "resource_group_name" {
    description = "Ressource group name"
    value = azurerm_resource_group.rg.name
}


output "MicrosoftAppId" {
    description = "Microsoft Application ID"
    value = azuread_application.aadappreg.id
}

output "BaseUrl" {
    description = "Microsoft Application ID"
    value = trimsuffix(azurerm_bot_channels_registration.teamsbotreg.endpoint, "/api/messages")
}

*/