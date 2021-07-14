# output "app_service" {
#   description = "Full description of the provisioned app service"
#   value       = azurerm_app_service.appservice
# }

output "appservice_base_url" {
  description = "Actual URL for our provisioned app service"
  value       = "https://${azurerm_app_service.appservice.default_site_hostname}"
}

output "existing_resource_group" {
  description = "Ressource group"
  value       = azurerm_resource_group.rg 
}

output "bot_channel_registration" {
  value = azurerm_bot_channels_registration.teamsbotreg.id
}

output "MicrosoftAppId" {
  description = "Microsoft Application ID"
  value       = azuread_application.aadappreg.application_id
}



