# Author::    Liam Bennett (mailto:liamjbennett@gmail.com)
# Copyright:: Copyright (c) 2014 Liam Bennett
# License::   MIT

# == Define msoffice::lip
#
# The define installs a language interface pack
#
# === Requirements/Dependencies
#
# Currently reequires the puppetlabs/stdlib module on the Puppet Forge in
# order to validate much of the the provided configuration.
#
# === Parameters
#
# [*version*]
# The version of office that was installed
#
# [*lang_code]
# The language code of the language to install
#
# [*arch*]
# The architecture version of office
#
# [*deployment_root*]
# The network location where the office installation media is stored
#
# [*deployment_root_absolute*]
# Whether the path to deployment_root is absolute (false by default)
#
# === Examples
#
#  Install the French Language Pack:
#
#    msoffice::lip {
#      version   => '2010',
#      lang_code => 'fr-fr',
#      arch      => 'x64'
#    }
#
define msoffice::lip(
  $version,
  $lang_code,
  $arch = 'x86',
  $deployment_root = '',
  $deployment_root_absolute = false,
) {

  include ::msoffice::params

  validate_re($version,'^(2003|2007|2010|2013)$', 'The version agrument specified does not match a valid version of office')
  validate_re($arch,'^(x86|x64)$', 'The arch argument specified does not match x86 or x64')

  $lang_regex = join(keys($msoffice::params::lcid_strings), '|')
  validate_re($lang_code,"^${lang_regex}$", 'The lang_code argument does not specifiy a valid language identifier')

  validate_bool($deployment_root_absolute)

  $office_num = $msoffice::params::office_versions[$version]['version']

  if ($deployment_root_absolute) {
    $lip_root = "${deployment_root}\\LIPs"
  } else {
    $lip_root = "${deployment_root}\\OFFICE${office_num}\\LIPs"
  }

  $lip_reg_key = "HKLM:\\SOFTWARE\\Microsoft\\Office\\${office_num}.0\\Common\\LanguageResources\\InstalledUIs"
  $lang_id = $msoffice::params::lcid_strings[$lang_code]
  $setup = "languageinterfacepack-${arch}-${lang_code}.exe"

  exec { 'install-lip':
    command   => "& \"${lip_root}\\${setup}\" /q /norestart",
    provider  => powershell,
    logoutput => true,
    onlyif    => "if (Get-Item -LiteralPath \'\\${lip_reg_key}\' -ErrorAction SilentlyContinue).GetValue(\'${lang_id}\')) { exit 1 }",
  }

}
