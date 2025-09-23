# Microsoft365DSC – October 2025 Major Release (version 1.25.1001.1)

As defined by our [Breaking Changes Policy](https://microsoft365dsc.com/concepts/breaking-changes/), twice a year we allow for breaking changes to be deployed as part of a release. Our next major release, scheduled to go out on October 21st 2025, will include several breaking changes and will be labeled version 1.25.1001.1. This article provides details on the breaking changes and other important updates that will be included as part of our October 2025 Major release.

## EXOGroupSettings - Renaming the UnifiedGroupWelcomeMessageEnabled Parameter ([#6531](https://github.com/microsoft/Microsoft365DSC/pull/6531))

The UnifiedGroupWelcomeMessageEnabled parameter of the EXOGroupsSettings resource will be renamed to WelcomeMessageEnabled to match the cmdlet's properties. To fix this in your existing configuration, simply search for instance of EXOGroupSettings and do a find and replace for the UnifiedGroupWelcomeMessageEnable property and replace it by WelcomeMessageEnabled.

## EXOHostedContentFilterPolicy - Removed Deprecated Parameters ([#6036](https://github.com/microsoft/Microsoft365DSC/pull/6036))

The following parameters were removed from the EXOHostedContentFIlterPolicy resource given they have been deprecated for some time already:

<ul>
<li>DownloadLink</li>
<li>EnableEndUserSpamNotifications</li>
<li>EndUserSpamNotificationCustomSubject</li>
<li>EndUserSpamNotificationFrequency</li>
<li>EndUserSpamNotificationLanguage</li>
</ul>

To fix existing configuration files, search for instance of these parameters in the configuration file and remove them.

## Removed 6 Intune Resources ([#6050](https://github.com/microsoft/Microsoft365DSC/pull/6050))

We've removed the following 6 resources due to them being deprecated or because of equivlent resource having been created (in the case of IntuneWifiConfigurationPolicyAndroidForWork):

<ul>
<li>IntuneDeviceCompliancePolicyAndroid</li>
<li>IntuneDeviceConfigurationPolicyAndroidDeviceAdministrator</li>
<li>IntuneTrustedRootCertificateAndroidEnterprise</li>
<li>IntuneVPNConfigurationPolicyAndroidEnterprise</li>
<li>IntuneWifiConfigurationPolicyAndroidDeviceAdministrator</li>
<li>IntuneWifiConfigurationPolicyAndroidForWork </li>
</ul>

To fix impacted configurations, simply remove any instances of the listed resources.

## EXORetentionPolicyTag - Changed AgeLimitForRetention from String to Integer ([#6103](https://github.com/microsoft/Microsoft365DSC/pull/6103))

The AgeLimitForRetention property of the EXORetentionPolicyTag was changed from being a String to being an Integer. To fix your configurations, simply search for this parameter and ensure they are specified as an Integer (e.g., remove quotes around the number).

## AADGroupEligibilitySchedule - Changed Mandatory Parameters ([#6124](https://github.com/microsoft/Microsoft365DSC/pull/6124))

In the AADGroupEligibilitySchedule, we've made the AccessId parameter mandatory and we've renamed the PrincipalDisplayName parameter to just principal and also made it mandatory. In order to fix your existing configurations, you need to make sure that every instance of AADGroupEligibilitySchedule specify both the Principal and AccessId parameters.

## IntuneAppProtectionPolicyAndroid and IntuneAppProtectionPolicyiOS Major Refactoring ([#6169](https://github.com/microsoft/Microsoft365DSC/pull/6169)), ([#6170](https://github.com/microsoft/Microsoft365DSC/pull/6170))

The IntuneAppProtectionPolicyAndroid and IntuneAppProtectionPolicyiOS resources underwent major refactoring as part of this major release. These changes were made to account for generic Intune assignments and to streamline its logic to better align it with other resources. Given the number of changes in the resources, it is recommended that existing instance of the resources in old configurations be reviewed manually to ensure alignment with the new changes. One alternative would be to use the Export/Snapshot feature to extract existing instances of these resources in the new proposed shape and include them as part of your configuration files.
