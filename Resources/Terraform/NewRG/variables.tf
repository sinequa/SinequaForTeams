variable "main_subscription_id" {
  type = string
}

variable "location" {
  type    = string
  default = "westeurope"
}

variable "rg_name" {
  description = "the name of the main resource group"
  type        = string
}

variable "kv_name" {
  description = "a unique name for the keyvault"
  type        = string
}

variable "aad_app_reg_name" {
  description = "the name of the aad app registration, must be unique on the tenant"
  type        = string
}

variable "bot_display_name" {
  description = "Friendly name for the bot up to 35 characters long"
  type        = string
}

variable "app_svc_plan" {
  description = "name of the app service plan"
  type        = string
}

variable "app_svc_name" {
  description = "a unique name for the app service instance, ASCII(7) letters from a to z, the digits from 0 to 9, and the hyphen (-), cannot start with a hyphen"
  type        = string
}


variable "env" {
  description = "Environment"
  default     = "dev"
  validation {
    condition     = contains(["dev", "", "prod"], var.env)
    error_message = "Allowed values for input_parameter are \"dev\", \"stage\", or \"prod\"."
  }
}

variable "app_reg_clientsecret" {
  description = "client secret ( bot framework authentication ): It must be at least 16 characters long, contain at least 1 upper or lower case alphabetical character, and contain at least 1 special character."
  type        = string
}

variable "sinequa_jwt" {
  description = "json web token - Sinequa Access Token"
  type        = string
}

variable "sinequa_hostname" {
  description = "Sinequa server FQDN"
  type        = string
}

variable "sinequa_appname" {
  description = "Sinequa SBA name"
  type        = string
}

variable "sinequa_wsqueryname" {
  description = "sinequa web service hostname"
  type        = string
}

variable "sinequa_domain" {
  description = "Sinequa security domain name (should contain an AAD partition)"
  type        = string
}

variable "sinequa_htts_port" {
  description = " Https port if non standard"
  type        = number
  default     = 443
}
