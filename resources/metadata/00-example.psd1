@{
  SourceFile = '01-telemetry.lgpo'
  Title = 'Telemetry and diagnostic data restraint'
  Owner = '<role-or-name>'
  IntroducedVersion = '2026-06-14-01'
  LastReviewed = '2026-06-14'

  IsoControls = @(
    'A.5.34 — Privacy and protection of PII'
    'A.8.16 — Monitoring activities (data minimisation)'
  )

  Justification = @'
Windows Diagnostic Data collection includes telemetry beyond what is required
for product functionality and security. Disabling reduces both the attack
surface for data exfiltration via Microsoft-approved channels and the volume
of PII transmitted to third parties, supporting the principle of data
minimisation under ISO 27001 A.5.34.

The AdvertisingInfo per-user policy is included because the advertising ID is
also a tracking identifier; disabling it aligns with the same justification.
'@

  UpstreamReferences = @(
    'https://learn.microsoft.com/en-us/windows/privacy/configure-windows-diagnostic-data-in-your-organization'
    'ADMX: DataCollection.admx, policy "AllowTelemetry"'
    'ADMX: ControlPanelDisplay.admx, policy "DisablePersonalisation"'
  )

  AppliesTo = @{
    WindowsEditions = @('Pro', 'Enterprise', 'Education')
    MinBuild = 19041
  }
}
