terraform {
  required_providers {
    azurerm = {
      source  = "hashicorp/azurerm"
      version = "=2.57.0"
    }
  }
}

module "naming" {
  source  = "Azure/naming/azurerm"
  suffix = [ "01" ] 
}


provider "azurerm" {
  partner_id = "947f5924-5e20-4f0a-96eb-808371995ac8" 
  features {}
  subscription_id = var.main_subscription_id  
}

locals {
   site_url = "https://${var.app_svc_name}.azurewebsites.net"
   }


data "azurerm_client_config" "current" {}


#existing resource group
data "azurerm_resource_group" "rg" {
  name    =  var.rg_name
}

#existing Key vault
data "azurerm_key_vault" "kv" {
  name                = var.kv_name
  resource_group_name = data.azurerm_resource_group.rg.name  
  # location            = azurerm_resource_group.rg.location
  # tenant_id           = data.azurerm_client_config.current.tenant_id
  # sku_name            = "standard"
}

#new app registration 
resource "azuread_application" "aadappreg" {
  display_name = var.aad_app_reg_name
  available_to_other_tenants = false
}

##App id and client secret  
resource "azuread_application_password" "aadappregpwd" {
  application_object_id = azuread_application.aadappreg.id
  description           = "My managed password"
  value                 = var.app_reg_clientsecret
  end_date              = "2099-01-01T01:02:03Z"
}


##bot channel registration 
resource "azurerm_bot_channels_registration" "teamsbotreg" {
  name                =  module.naming.bot_channel_ms_teams.name_unique  #"bcr-${var.aad_app_reg_name}"
  display_name        =  var.bot_display_name
  location            = "global"
  resource_group_name = data.azurerm_resource_group.rg.name
  sku                 = "F0"
  microsoft_app_id    = azuread_application.aadappreg.application_id
  endpoint            = "${local.site_url}/api/messages"
  tags = {
    environment = var.env
  }
}


##bot teams channel 
resource "azurerm_bot_channel_ms_teams" "teamschannel" {
  bot_name            = azurerm_bot_channels_registration.teamsbotreg.name
  location            = azurerm_bot_channels_registration.teamsbotreg.location
  resource_group_name = data.azurerm_resource_group.rg.name
}


#app service plan --- Could be an ASE and could be mutualized
data "azurerm_app_service_plan" "appserviceplan" {
  name = var.app_svc_plan
  resource_group_name = data.azurerm_resource_group.rg.name 
}


#create a new app service using this plan
resource "azurerm_app_service" "appservice" {

  name = var.app_svc_name
  resource_group_name = data.azurerm_resource_group.rg.name 
  app_service_plan_id = data.azurerm_app_service_plan.appserviceplan.application_id 
  location = var.location
  site_config {
    use_32_bit_worker_process = true
    dotnet_framework_version = "v5.0"
    scm_type                 = "None" #TODO, git ? VS ?  
  }
  #system assigned identity ( key vault access )
  identity  {
    type = "SystemAssigned"
  }
  #1st 3 entries are keyvault references . 
  app_settings = {
    "MicrosoftAppId" = "@Microsoft.KeyVault(VaultName=${var.kv_name};SecretName=msftappid)"
    "MicrosoftAppPassword" = "@Microsoft.KeyVault(VaultName=${var.kv_name};SecretName=msftapppwd)"
    "JWT" = "@Microsoft.KeyVault(VaultName=${var.kv_name};SecretName=jwt)"
    "Sinequa:BaseUrl" = local.site_url #"trimsuffix(azurerm_bot_channels_registration.teamsbotreg.endpoint, "/api/messages")
    "Sinequa:Domain" = var.sinequa_domain
    "Sinequa:HostName" = var.sinequa_hostname
    "Sinequa:AppName" = var.sinequa_appname
    "Sinequa:WSQueryName" = var.sinequa_wsqueryname
    "Sinequa:CustomPort" = var.sinequa_htts_port
  }
  #depends_on = [azurerm_key_vault_access_policy.appservicepolicy] 
  tags = {
    environment = var.env
  }
}


#**** useless since we are using default azure hostnames for our app service.****  
# resource "azurerm_app_service_custom_hostname_binding" "hostname_binding" {
#   hostname            = "${azurerm_app_service.appservice.name}.azurewebsites.net"
#   app_service_name    = azurerm_app_service.appservice.name
#   resource_group_name = azurerm_resource_group.rg.name
# }





#identity based policy
resource "azurerm_key_vault_access_policy" "appservicepolicy" {
  key_vault_id = data.azurerm_key_vault.kv.id
  tenant_id    = data.azurerm_client_config.current.tenant_id
  object_id    =  azurerm_app_service.appservice.identity.0.principal_id 
  
  secret_permissions = [
    "Get"
  ]
}

# resource "azurerm_key_vault_access_policy" "service" {
#   key_vault_id = azurerm_key_vault.kv.id

#   tenant_id = data.azurerm_client_config.current.tenant_id
#   object_id = data.azurerm_client_config.current.object_id

#   key_permissions = [
#   ]

#   secret_permissions = [
#     "Get",
#     "Set",
#     "List",
#     "Delete",
#   ]
# }

#adding all required secrets
resource "azurerm_key_vault_secret" "appid" {
  name         = "msftappid"
  value        = azuread_application.aadappreg.application_id
  key_vault_id = data.azurerm_key_vault.kv.id
  # depends_on = [azurerm_key_vault_access_policy.service]
}
resource "azurerm_key_vault_secret" "apppwd" {
  name         = "msftapppwd"
  value        = var.app_reg_clientsecret
  key_vault_id = data.azurerm_key_vault.kv.id
  # depends_on = [azurerm_key_vault_access_policy.service]
}
resource "azurerm_key_vault_secret" "jwt" {
  name         = "jwt"
   value        = var.sinequa_jwt
  key_vault_id =  data.azurerm_key_vault.kv.id
  # depends_on = [azurerm_key_vault_access_policy.service]
}


