# Microsoft365DSC – October 2025 Major Release (version 1.25.1001.1)

As defined by our [Breaking Changes Policy](https://microsoft365dsc.com/concepts/breaking-changes/), twice a year we allow for breaking changes to be deployed as part of a release. Our next major release, scheduled to go out on October 21st 2025, will include several breaking changes and will be labeled version 1.25.1001.1. This article provides details on the breaking changes and other important updates that will be included as part of our October 2025 Major release.

## EXOGroupSettings - Renaming the UnifiedGroupWelcomeMessageEnabled Parameter  ([#6531](https://github.com/microsoft/Microsoft365DSC/pull/6531))

The UnifiedGroupWelcomeMessageEnabled parameter of the EXOGroupsSettings resource will be renamed to WelcomeMessageEnabled to match the cmdlet's properties. To fix this in your existing configuration, simply search for instance of EXOGroupSettings and do a find and replace for the UnifiedGroupWelcomeMessageEnable property and replace it by WelcomeMessageEnabled.

## EXOHostedContentFilterPolicy - Removed Deprecated Parameters

The following parameters were removed from the EXOHostedContentFIlterPolicy resource given they have been deprecated for some time already:

<ul>
<li>DownloadLink</li>
<li>EnableEndUserSpamNotifications</li>
<li>EndUserSpamNotificationCustomSubject</li>
<li>EndUserSpamNotificationFrequency</li>
<li>EndUserSpamNotificationLanguage</li>
</ul>

To fix existing configuration files, search for instance of these parameters in the configuration file and remove them.

## Removed 6 Intune Resources

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
