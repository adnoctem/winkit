@{
  FieldMaps = @{
    LogonType = @{
      2 = @{
        Name = 'Interactive'
        Description = 'Local keyboard/screen logon.'
        Examples = @('Console sign-in', 'Physical workstation logon')
      }
      3 = @{
        Name = 'Network'
        Description = 'Network logon without interactive session.'
        Examples = @('SMB access', 'Remote file access', 'Printer/share access')
      }
      4 = @{
        Name = 'Batch'
        Description = 'Batch logon.'
        Examples = @('Scheduled task using stored credentials')
      }
      5 = @{
        Name = 'Service'
        Description = 'Service Control Manager started a service account logon session.'
        Examples = @('Windows service account startup')
      }
      7 = @{
        Name = 'Unlock'
        Description = 'Existing workstation session was unlocked.'
        Examples = @('User unlocks workstation')
      }
      8 = @{
        Name = 'NetworkCleartext'
        Description = 'Network logon where credentials were passed in cleartext to the authentication package.'
        Examples = @('IIS Basic authentication in some configurations')
        Risk = 'Review carefully; cleartext credential handling may be involved.'
      }
      9 = @{
        Name = 'NewCredentials'
        Description = 'Local identity reused with alternate outbound network credentials.'
        Examples = @('runas /netonly')
      }
      10 = @{
        Name = 'RemoteInteractive'
        Description = 'Remote interactive logon.'
        Examples = @('RDP', 'Remote Desktop Services')
      }
      11 = @{
        Name = 'CachedInteractive'
        Description = 'Interactive logon using cached domain credentials.'
        Examples = @('Laptop domain user logon while DC is unreachable')
      }
    }

    ImpersonationLevel = @{
      '%%1832' = @{
        Name = 'Anonymous'
        Description = 'Server process cannot identify the client.'
      }
      '%%1833' = @{
        Name = 'Identification'
        Description = 'Server can identify the client but cannot impersonate it.'
      }
      '%%1834' = @{
        Name = 'Impersonation'
        Description = 'Server can impersonate the client on the local system.'
      }
      '%%1835' = @{
        Name = 'Delegation'
        Description = 'Server can impersonate the client on remote systems.'
      }
    }

    AuditResult = @{
      Success = @{ Name = 'Success' }
      Failure = @{ Name = 'Failure' }
    }
  }

  Groups = @{
    Logon = @{
      Name = 'Logon'
      Description = 'Interactive, network, service, batch, RDP, Kerberos, NTLM, lock/unlock and related authentication events.'
      DefaultLogs = @('Security')
      EventIds = @(
        4624, 4625, 4634, 4647, 4648, 4672, 4740, 4767,
        4768, 4769, 4771, 4776, 4800, 4801
      )
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName',
        'TargetUserName', 'TargetDomainName', 'SubjectUserName', 'SubjectDomainName',
        'LogonType', 'LogonTypeName', 'IpAddress', 'WorkstationName',
        'AuthenticationPackageName', 'LogonProcessName', 'ProcessName',
        'TargetLogonId', 'Status', 'SubStatus'
      )
      DefaultSuppressions = @{
        TargetUserName = @('ANONYMOUS LOGON', 'LOCAL SERVICE', 'NETWORK SERVICE', 'SYSTEM')
        TargetUserNameEndsWith = @('$')
      }
    }

    AccountChange = @{
      Name = 'AccountChange'
      Description = 'User and group account lifecycle and membership changes.'
      DefaultLogs = @('Security')
      EventIds = @(
        4720, 4722, 4723, 4724, 4725, 4726, 4727, 4728, 4729,
        4730, 4731, 4732, 4733, 4734, 4738, 4756, 4757, 4767
      )
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName',
        'TargetUserName', 'TargetDomainName', 'SubjectUserName', 'SubjectDomainName'
      )
    }

    Service = @{
      Name = 'Service'
      Description = 'Service lifecycle, service failures, and service installation events.'
      DefaultLogs = @('System', 'Security')
      EventIds = @(4697, 7000, 7001, 7031, 7034, 7036, 7040, 7045)
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName',
        'ServiceName', 'ServiceFileName', 'ImagePath'
      )
    }

    BootShutdown = @{
      Name = 'BootShutdown'
      Description = 'Boot, shutdown, crash, and unexpected reboot events.'
      DefaultLogs = @('System')
      EventIds = @(41, 1001, 1074, 6005, 6006, 6008)
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName'
      )
    }

    ApplicationCrash = @{
      Name = 'ApplicationCrash'
      Description = 'Application crashes, unhandled exceptions, and Windows Error Reporting events.'
      DefaultLogs = @('Application')
      EventIds = @(1000, 1001, 1026)
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName'
      )
    }

    Installer = @{
      Name = 'Installer'
      Description = 'MSI installer product installation and removal events.'
      DefaultLogs = @('Application')
      EventIds = @(11707, 11708, 11724)
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName'
      )
    }

    PowerShell = @{
      Name = 'PowerShell'
      Description = 'PowerShell engine, module, and script block logging events.'
      DefaultLogs = @('Microsoft-Windows-PowerShell/Operational')
      EventIds = @(400, 403, 600, 800, 4103, 4104, 4105, 4106)
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName',
        'ScriptBlockText', 'ScriptBlockId', 'UserName'
      )
    }

    ScheduledTask = @{
      Name = 'ScheduledTask'
      Description = 'Scheduled task registration, execution, action, update and deletion events.'
      DefaultLogs = @('Microsoft-Windows-TaskScheduler/Operational')
      EventIds = @(
        100, 101, 102, 103, 106, 107, 108, 110, 111,
        118, 119, 129, 140, 141, 142, 200, 201, 203
      )
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName',
        'TaskName', 'TaskContent'
      )
    }

    Sysmon = @{
      Name = 'Sysmon'
      Description = 'Optional Sysmon process, network, file, registry, WMI, DNS, pipe and tampering telemetry.'
      DefaultLogs = @('Microsoft-Windows-Sysmon/Operational')
      EventIds = @(
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
        16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 255
      )
      Optional = $true
      ImportantFields = @(
        'TimeCreated', 'Id', 'ProviderName', 'MachineName',
        'Image', 'CommandLine', 'ProcessGuid', 'ProcessId',
        'SourceIp', 'DestinationIp', 'SourcePort', 'DestinationPort'
      )
    }
  }

  Events = @{
    Security = @{
      Logon = @(
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4624
          Name = 'SuccessfulLogon'
          DisplayName = 'An account was successfully logged on'
          Category = 'Logon'
          Subcategory = 'Audit Logon'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the account that requested the logon.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account name that requested the logon.' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain or computer name of the subject account.' }
            SubjectLogonId = @{ XmlName = 'SubjectLogonId'; Type = 'HexInt64'; Description = 'Logon ID of the subject account.' }
            TargetUserSid = @{ XmlName = 'TargetUserSid'; Type = 'String'; Description = 'SID of the account that logged on.' }
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name that logged on.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain or computer name of the account that logged on.' }
            TargetLogonId = @{ XmlName = 'TargetLogonId'; Type = 'HexInt64'; Description = 'Logon ID assigned to the new logon session.' }
            LogonType = @{ XmlName = 'LogonType'; Type = 'Int32'; Map = 'LogonType'; Description = 'Numeric logon type. Critical for interpreting 4624.' }
            LogonProcessName = @{ XmlName = 'LogonProcessName'; Type = 'String'; Description = 'Trusted logon process used for the logon.' }
            AuthenticationPackageName = @{ XmlName = 'AuthenticationPackageName'; Type = 'String'; Description = 'Authentication package used (Kerberos, NTLM, Negotiate).' }
            WorkstationName = @{ XmlName = 'WorkstationName'; Type = 'String'; Description = 'Workstation name supplied by the client where available.' }
            IpAddress = @{ XmlName = 'IpAddress'; Type = 'String'; Description = 'Source IP address where available.' }
            IpPort = @{ XmlName = 'IpPort'; Type = 'String'; Description = 'Source port where available.' }
            ProcessId = @{ XmlName = 'ProcessId'; Type = 'HexInt64'; Description = 'Process ID of the process that requested the logon.' }
            ProcessName = @{ XmlName = 'ProcessName'; Type = 'String'; Description = 'Process path of the process that requested the logon.' }
            ImpersonationLevel = @{ XmlName = 'ImpersonationLevel'; Type = 'String'; Map = 'ImpersonationLevel'; Description = 'Impersonation level assigned to the logon session.' }
            ElevatedToken = @{ XmlName = 'ElevatedToken'; Type = 'String'; Description = 'Indicates whether the session uses an elevated token where present.' }
            RestrictedAdminMode = @{ XmlName = 'RestrictedAdminMode'; Type = 'String'; Description = 'RDP Restricted Admin mode state where present.' }
            TargetOutboundUserName = @{ XmlName = 'TargetOutboundUserName'; Type = 'String'; Description = 'Outbound identity for NewCredentials logons where present.' }
            TargetOutboundDomainName = @{ XmlName = 'TargetOutboundDomainName'; Type = 'String'; Description = 'Outbound domain for NewCredentials logons where present.' }
          }
          Notes = @(
            'LogonType is the primary discriminator for useful interpretation.',
            'LogonType 3 is common and noisy on file servers, print servers, and domain-joined systems.',
            'LogonType 10 is typically RDP/Remote Desktop.',
            'Correlate TargetLogonId with later 4634/4647 logoff events on the same host.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4624',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4624'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4625
          Name = 'FailedLogon'
          DisplayName = 'An account failed to log on'
          Category = 'Logon'
          Subcategory = 'Audit Logon'
          Result = 'Failure'
          UtilityGroup = 'Logon'
          Fields = @{
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the account that requested the logon.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account name that requested the logon.' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain or computer name of the subject account.' }
            SubjectLogonId = @{ XmlName = 'SubjectLogonId'; Type = 'HexInt64'; Description = 'Logon ID of the subject account.' }
            TargetUserSid = @{ XmlName = 'TargetUserSid'; Type = 'String'; Description = 'SID of the account that failed to log on.' }
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name that failed to log on.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain or computer name of the account that failed to log on.' }
            Status = @{ XmlName = 'Status'; Type = 'String'; Description = 'NTSTATUS code indicating the failure reason.' }
            SubStatus = @{ XmlName = 'SubStatus'; Type = 'String'; Description = 'Sub-status code providing additional failure detail.' }
            LogonType = @{ XmlName = 'LogonType'; Type = 'Int32'; Map = 'LogonType'; Description = 'Numeric logon type. Critical for interpreting 4625.' }
            LogonProcessName = @{ XmlName = 'LogonProcessName'; Type = 'String'; Description = 'Trusted logon process used for the logon attempt.' }
            AuthenticationPackageName = @{ XmlName = 'AuthenticationPackageName'; Type = 'String'; Description = 'Authentication package used for the attempt.' }
            WorkstationName = @{ XmlName = 'WorkstationName'; Type = 'String'; Description = 'Workstation name supplied by the client where available.' }
            IpAddress = @{ XmlName = 'IpAddress'; Type = 'String'; Description = 'Source IP address where available.' }
            IpPort = @{ XmlName = 'IpPort'; Type = 'String'; Description = 'Source port where available.' }
            ProcessId = @{ XmlName = 'ProcessId'; Type = 'HexInt64'; Description = 'Process ID that requested the logon attempt.' }
            ProcessName = @{ XmlName = 'ProcessName'; Type = 'String'; Description = 'Process path that requested the logon attempt.' }
            FailureReason = @{ XmlName = 'FailureReason'; Type = 'String'; Description = 'Translated failure reason text.' }
          }
          Notes = @(
            'Common Status/SubStatus values: 0xC0000064 (no such user), 0xC000006A (bad password), 0xC0000234 (account locked), 0xC0000072 (account disabled).',
            'Repeated 4625 events from a single source IP followed by a 4624 Success may indicate password guessing.',
            'LogonType 3 failures are common; LogonType 10 failures from unexpected sources warrant investigation.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4625',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4625'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4634
          Name = 'AccountLogoff'
          DisplayName = 'An account was logged off'
          Category = 'Logon'
          Subcategory = 'Audit Logoff'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserSid = @{ XmlName = 'TargetUserSid'; Type = 'String'; Description = 'SID of the account that was logged off.' }
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name that was logged off.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the account that was logged off.' }
            TargetLogonId = @{ XmlName = 'TargetLogonId'; Type = 'HexInt64'; Description = 'Logon ID of the session that was ended.' }
            LogonType = @{ XmlName = 'LogonType'; Type = 'Int32'; Map = 'LogonType'; Description = 'Numeric logon type of the ended session.' }
          }
          Notes = @(
            'Correlate TargetLogonId with 4624 logon events on the same machine.',
            'LogonType 3 sessions are often logged off quickly; long-lived session logoffs may indicate interactive activity.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4634',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4634'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4647
          Name = 'UserInitiatedLogoff'
          DisplayName = 'User initiated logoff'
          Category = 'Logon'
          Subcategory = 'Audit Logoff'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserSid = @{ XmlName = 'TargetUserSid'; Type = 'String'; Description = 'SID of the user who initiated the logoff.' }
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'User who initiated the logoff.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the user.' }
            TargetLogonId = @{ XmlName = 'TargetLogonId'; Type = 'HexInt64'; Description = 'Logon ID of the session that was ended.' }
          }
          Notes = @(
            'User-initiated logoff (Start menu sign out, shutdown, etc.) vs system-initiated logoff (4634).',
            'Correlate with 4624 to measure session duration.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4647',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4647'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4648
          Name = 'ExplicitCredentialLogon'
          DisplayName = 'A logon was attempted using explicit credentials'
          Category = 'Logon'
          Subcategory = 'Audit Logon'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the account requesting the explicit credential logon.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account that initiated the explicit credential logon.' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain of the subject account.' }
            SubjectLogonId = @{ XmlName = 'SubjectLogonId'; Type = 'HexInt64'; Description = 'Existing logon session that requested the alternate credentials.' }
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name whose credentials were used.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the target account.' }
            TargetServerName = @{ XmlName = 'TargetServerName'; Type = 'String'; Description = 'Server or service the alternate credentials were targeted at.' }
            ProcessId = @{ XmlName = 'ProcessId'; Type = 'HexInt64'; Description = 'Process ID that requested the explicit credentials.' }
            ProcessName = @{ XmlName = 'ProcessName'; Type = 'String'; Description = 'Process path that requested the explicit credentials.' }
            IpAddress = @{ XmlName = 'IpAddress'; Type = 'String'; Description = 'Source IP address where available.' }
          }
          Notes = @(
            'Often produced by runas, credential manager use, or lateral movement tools.',
            'SubjectUserSid != TargetUserSid indicates identity switching.',
            'ProcessName can reveal the tool performing the credential handoff.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4648',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4648'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4672
          Name = 'SpecialPrivilegesAssigned'
          DisplayName = 'Special privileges assigned to new logon'
          Category = 'Logon'
          Subcategory = 'Audit Sensitive Privilege Use'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the account to which privileges were assigned.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account name to which privileges were assigned.' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain of the account.' }
            SubjectLogonId = @{ XmlName = 'SubjectLogonId'; Type = 'HexInt64'; Description = 'Logon ID of the session receiving the privileges.' }
            PrivilegeList = @{ XmlName = 'PrivilegeList'; Type = 'String'; Description = 'Special privileges assigned to the logon session.' }
          }
          Notes = @(
            'E.g. SeSecurityPrivilege, SeBackupPrivilege, SeDebugPrivilege, SeTcbPrivilege.',
            'SeDebugPrivilege is often assigned to administrative accounts but can also indicate privilege escalation.',
            'Correlate with 4624 events sharing the same TargetLogonId/SubjectLogonId.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4672',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4672'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4740
          Name = 'AccountLocked'
          DisplayName = 'A user account was locked out'
          Category = 'Logon'
          Subcategory = 'Audit User Account Management'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserSid = @{ XmlName = 'TargetUserSid'; Type = 'String'; Description = 'SID of the account that was locked out.' }
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name that was locked out.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the locked account.' }
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the workstation that caused the lockout.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account name that caused the lockout (usually a machine account).' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain of the subject.' }
          }
          Notes = @(
            'Check 4625 events with the same TargetUserName for the source IP/workstation that caused the lockout.',
            'DC-only event for domain accounts.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4740',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4740'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4767
          Name = 'AccountUnlocked'
          DisplayName = 'A user account was unlocked'
          Category = 'Logon'
          Subcategory = 'Audit User Account Management'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Notes = @('Account was unlocked, either manually by an administrator or automatically by lockout policy expiration.')
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4767',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4767'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4768
          Name = 'KerberosTgtRequested'
          DisplayName = 'A Kerberos authentication ticket (TGT) was requested'
          Category = 'Logon'
          Subcategory = 'Audit Kerberos Authentication Service'
          Result = '$undefined'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account that requested the TGT.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the account.' }
            TargetSid = @{ XmlName = 'TargetSid'; Type = 'String'; Description = 'SID of the account.' }
            ServiceName = @{ XmlName = 'ServiceName'; Type = 'String'; Description = 'Kerberos service principal name.' }
            IpAddress = @{ XmlName = 'IpAddress'; Type = 'String'; Description = 'Source IP address of the TGT request.' }
            IpPort = @{ XmlName = 'IpPort'; Type = 'String'; Description = 'Source port of the TGT request.' }
            TicketEncryptionType = @{ XmlName = 'TicketEncryptionType'; Type = 'String'; Description = 'Encryption type used for the ticket (e.g. 0x12 = AES256).' }
            PreAuthType = @{ XmlName = 'PreAuthType'; Type = 'String'; Description = 'Pre-authentication type (e.g. 2 = password, 15 = smart card).' }
            Status = @{ XmlName = 'Status'; Type = 'String'; Description = 'NTSTATUS code for the result.' }
            CertIssuerName = @{ XmlName = 'CertIssuerName'; Type = 'String'; Description = 'Certificate issuer for smart card logons where present.' }
            CertSerialNumber = @{ XmlName = 'CertSerialNumber'; Type = 'String'; Description = 'Certificate serial number for smart card logons where present.' }
            CertThumbprint = @{ XmlName = 'CertThumbprint'; Type = 'String'; Description = 'Certificate thumbprint for smart card logons where present.' }
          }
          Notes = @(
            'DC-only event. Generated on the KDC when a TGT is requested.',
            'Encryption type downgrade (e.g. RC4 instead of AES) can indicate Kerberoasting or Golden Ticket activity.',
            'Correlate with 4769 (service ticket) and 4771 (pre-auth failure) for the same user.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4768',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4768'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4769
          Name = 'KerberosServiceTicket'
          DisplayName = 'A Kerberos service ticket was requested'
          Category = 'Logon'
          Subcategory = 'Audit Kerberos Service Ticket Operations'
          Result = '$undefined'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account that requested the service ticket.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the account.' }
            ServiceName = @{ XmlName = 'ServiceName'; Type = 'String'; Description = 'Kerberos service principal name of the target service.' }
            ServiceSid = @{ XmlName = 'ServiceSid'; Type = 'String'; Description = 'SID of the service account.' }
            TicketOptions = @{ XmlName = 'TicketOptions'; Type = 'String'; Description = 'Bitmask of Kerberos ticket options.' }
            TicketEncryptionType = @{ XmlName = 'TicketEncryptionType'; Type = 'String'; Description = 'Encryption type used for the service ticket.' }
            IpAddress = @{ XmlName = 'IpAddress'; Type = 'String'; Description = 'Source IP address of the service ticket request.' }
            IpPort = @{ XmlName = 'IpPort'; Type = 'String'; Description = 'Source port of the service ticket request.' }
            Status = @{ XmlName = 'Status'; Type = 'String'; Description = 'NTSTATUS code for the result.' }
          }
          Notes = @(
            'DC-only event. Generated when a service ticket is requested for a specific service.',
            'TicketEncryptionType of 0x17 (RC4) without a corresponding 4768 TGT with the same type may indicate a Silver Ticket.',
            'Failed attempts (Status != 0x0) may indicate Kerberoasting attempts.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4769',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4769'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4771
          Name = 'KerberosPreAuthFailed'
          DisplayName = 'Kerberos pre-authentication failed'
          Category = 'Logon'
          Subcategory = 'Audit Kerberos Authentication Service'
          Result = 'Failure'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name that failed pre-authentication.' }
            TargetSid = @{ XmlName = 'TargetSid'; Type = 'String'; Description = 'SID of the account.' }
            ServiceName = @{ XmlName = 'ServiceName'; Type = 'String'; Description = 'Kerberos service principal name requested.' }
            IpAddress = @{ XmlName = 'IpAddress'; Type = 'String'; Description = 'Source IP address of the failed request.' }
            IpPort = @{ XmlName = 'IpPort'; Type = 'String'; Description = 'Source port of the failed request.' }
            Status = @{ XmlName = 'Status'; Type = 'String'; Description = 'NTSTATUS code (0x18 = bad password, 0x6 = unknown user).' }
            PreAuthType = @{ XmlName = 'PreAuthType'; Type = 'String'; Description = 'Pre-authentication type used.' }
          }
          Notes = @(
            'DC-only event. Status 0x18 = wrong password; Status 0x6 = account does not exist.',
            'Multiple failures followed by success indicate password guessing/brute force.',
            'AS-REP Roasting is observable when PreAuthType is 0 (no pre-authentication).'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4771',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4771'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4776
          Name = 'NtlmCredentialValidation'
          DisplayName = 'The domain controller attempted to validate the credentials for an account'
          Category = 'Logon'
          Subcategory = 'Audit Credential Validation'
          Result = '$undefined'
          UtilityGroup = 'Logon'
          Fields = @{
            TargetUserName = @{ XmlName = 'TargetUserName'; Type = 'String'; Description = 'Account name validated.' }
            TargetDomainName = @{ XmlName = 'TargetDomainName'; Type = 'String'; Description = 'Domain of the account.' }
            WorkstationName = @{ XmlName = 'WorkstationName'; Type = 'String'; Description = 'Workstation from which the logon attempt originated.' }
            Status = @{ XmlName = 'Status'; Type = 'String'; Description = 'NTSTATUS code for the validation result.' }
          }
          Notes = @(
            'DC-only event. Success = 0x0, bad password = 0xC000006A.',
            'NTLM authentication may indicate downgrade attacks or legacy protocol usage.',
            'Multiple failures from the same workstation for different users may indicate credential dumping/harvesting.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4776',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4776'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4800
          Name = 'WorkstationLocked'
          DisplayName = 'The workstation was locked'
          Category = 'Logon'
          Subcategory = 'Audit Other Logon/Logoff Events'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Notes = @('User locked the workstation (Windows+L).')
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4800',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4800'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4801
          Name = 'WorkstationUnlocked'
          DisplayName = 'The workstation was unlocked'
          Category = 'Logon'
          Subcategory = 'Audit Other Logon/Logoff Events'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Notes = @('User unlocked the workstation.')
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4801',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4801'
          )
        }
      )

      ProcessExecution = @(
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4688
          Name = 'ProcessCreated'
          DisplayName = 'A new process has been created'
          Category = 'Process Execution'
          Subcategory = 'Audit Process Creation'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Fields = @{
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the account that created the process.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account that created the process.' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain of the subject account.' }
            SubjectLogonId = @{ XmlName = 'SubjectLogonId'; Type = 'HexInt64'; Description = 'Logon ID of the subject.' }
            NewProcessId = @{ XmlName = 'NewProcessId'; Type = 'HexInt64'; Description = 'Hex PID of the created process.' }
            NewProcessName = @{ XmlName = 'NewProcessName'; Type = 'String'; Description = 'Full path of the created process executable.' }
            TokenElevationType = @{ XmlName = 'TokenElevationType'; Type = 'String'; Description = 'Elevation type of the process token (%%1936-%%1938).' }
            ProcessId = @{ XmlName = 'ProcessId'; Type = 'HexInt64'; Description = 'PID of the parent process.' }
            CommandLine = @{ XmlName = 'CommandLine'; Type = 'String'; Description = 'Full command line of the created process.' }
            ParentProcessName = @{ XmlName = 'ParentProcessName'; Type = 'String'; Description = 'Path of the parent process executable.' }
            MandatoryLabel = @{ XmlName = 'MandatoryLabel'; Type = 'String'; Description = 'Integrity level of the new process.' }
          }
          Notes = @(
            'Requires Audit Process Creation policy to be enabled (not on by default).',
            'Command line is available from Windows 8.1/Server 2012 R2 onwards.',
            'ParentProcessName requires Windows 10/Server 2016 or newer.',
            'Unexpected parent processes (e.g. Word spawning PowerShell) warrant investigation.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4688',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4688'
          )
        }
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4689
          Name = 'ProcessExited'
          DisplayName = 'A process has exited'
          Category = 'Process Execution'
          Subcategory = 'Audit Process Termination'
          Result = 'Success'
          UtilityGroup = 'Logon'
          Notes = @('Correlate ProcessId with 4688 to compute process runtime.')
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4689',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4689'
          )
        }
      )

      AccountChange = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4720; Name = 'UserAccountCreated'; DisplayName = 'A user account was created'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; Notes = @('CreatedBy reveals which admin performed the creation.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4720', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4720') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4722; Name = 'UserAccountEnabled'; DisplayName = 'A user account was enabled'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4722', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4722') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4723; Name = 'PasswordChangeAttempt'; DisplayName = 'An attempt was made to change an account password'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = '$undefined'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4723', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4723') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4724; Name = 'PasswordResetAttempt'; DisplayName = 'An attempt was made to reset an account password'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; Notes = @('Password reset by administrator, not self-service change.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4724', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4724') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4725; Name = 'UserAccountDisabled'; DisplayName = 'A user account was disabled'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4725', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4725') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4726; Name = 'UserAccountDeleted'; DisplayName = 'A user account was deleted'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4726', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4726') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4738; Name = 'UserAccountChanged'; DisplayName = 'A user account was changed'; Category = 'Account Change'; Subcategory = 'Audit User Account Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; Notes = @('Attributes changed: display name, home directory, profile path, script path, password last set, etc.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4738', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4738') }
      )

      GroupChange = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4727; Name = 'GlobalGroupCreated'; DisplayName = 'A security-enabled global group was created'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4727', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4727') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4728; Name = 'MemberAddedToGlobalGroup'; DisplayName = 'A member was added to a security-enabled global group'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; Notes = @('Check for additions to sensitive groups (Domain Admins, Enterprise Admins, etc.).'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4728', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4728') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4729; Name = 'MemberRemovedFromGlobalGroup'; DisplayName = 'A member was removed from a security-enabled global group'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4729', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4729') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4730; Name = 'GlobalGroupDeleted'; DisplayName = 'A security-enabled global group was deleted'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4730', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4730') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4731; Name = 'LocalGroupCreated'; DisplayName = 'A security-enabled local group was created'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4731', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4731') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4732; Name = 'MemberAddedToLocalGroup'; DisplayName = 'A member was added to a security-enabled local group'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; Notes = @('Check for additions to Administrators, Remote Desktop Users, etc.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4732', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4732') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4733; Name = 'MemberRemovedFromLocalGroup'; DisplayName = 'A member was removed from a security-enabled local group'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4733', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4733') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4734; Name = 'LocalGroupDeleted'; DisplayName = 'A security-enabled local group was deleted'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4734', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4734') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4756; Name = 'MemberAddedToUniversalGroup'; DisplayName = 'A member was added to a security-enabled universal group'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4756', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4756') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4757; Name = 'MemberRemovedFromUniversalGroup'; DisplayName = 'A member was removed from a security-enabled universal group'; Category = 'Group Change'; Subcategory = 'Audit Security Group Management'; Result = 'Success'; UtilityGroup = 'AccountChange'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4757', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4757') }
      )

      Service = @(
        @{
          LogName = 'Security'
          ProviderName = 'Microsoft-Windows-Security-Auditing'
          Id = 4697
          Name = 'ServiceInstalled'
          DisplayName = 'A service was installed in the system'
          Category = 'Service Installation'
          Subcategory = 'Audit Security System Extension'
          Result = 'Success'
          UtilityGroup = 'Service'
          Fields = @{
            SubjectUserSid = @{ XmlName = 'SubjectUserSid'; Type = 'String'; Description = 'SID of the account that installed the service.' }
            SubjectUserName = @{ XmlName = 'SubjectUserName'; Type = 'String'; Description = 'Account that installed the service.' }
            SubjectDomainName = @{ XmlName = 'SubjectDomainName'; Type = 'String'; Description = 'Domain of the account.' }
            ServiceName = @{ XmlName = 'ServiceName'; Type = 'String'; Description = 'Name of the installed service.' }
            ServiceFileName = @{ XmlName = 'ServiceFileName'; Type = 'String'; Description = 'Service binary path (e.g. exe, dll).' }
            ServiceType = @{ XmlName = 'ServiceType'; Type = 'String'; Description = 'Service type (e.g. Kernel Driver, File System Driver).' }
            ServiceStartType = @{ XmlName = 'ServiceStartType'; Type = 'String'; Description = 'Service start type (Auto, Demand, Disabled).' }
            ServiceAccount = @{ XmlName = 'ServiceAccount'; Type = 'String'; Description = 'Account the service runs as (LocalSystem, NetworkService, etc.).' }
          }
          Notes = @(
            'Requires Audit Security System Extension policy to be enabled.',
            'Service installation is a common persistence mechanism.',
            'Correlate with System log event 7045 which also indicates service installation.'
          )
          References = @(
            'https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4697',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4697'
          )
        }
      )

      ScheduledTask = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4698; Name = 'ScheduledTaskCreated'; DisplayName = 'A scheduled task was created'; Category = 'Scheduled Task'; Subcategory = 'Audit Other Object Access Events'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; Notes = @('Check TaskContent XML for suspicious commands, scripts, or persistence indicators.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4698', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4698') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4699; Name = 'ScheduledTaskDeleted'; DisplayName = 'A scheduled task was deleted'; Category = 'Scheduled Task'; Subcategory = 'Audit Other Object Access Events'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4699', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4699') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4700; Name = 'ScheduledTaskEnabled'; DisplayName = 'A scheduled task was enabled'; Category = 'Scheduled Task'; Subcategory = 'Audit Other Object Access Events'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4700', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4700') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4701; Name = 'ScheduledTaskDisabled'; DisplayName = 'A scheduled task was disabled'; Category = 'Scheduled Task'; Subcategory = 'Audit Other Object Access Events'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4701', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4701') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4702; Name = 'ScheduledTaskUpdated'; DisplayName = 'A scheduled task was updated'; Category = 'Scheduled Task'; Subcategory = 'Audit Other Object Access Events'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4702', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4702') }
      )

      AuditPolicy = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4719; Name = 'AuditPolicyChanged'; DisplayName = 'System audit policy was changed'; Category = 'Audit Policy'; Subcategory = 'Audit Audit Policy Change'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Audit policy changes may indicate attempts to disable logging before malicious activity.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4719', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4719') }
      )

      Tampering = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 1102; Name = 'AuditLogCleared'; DisplayName = 'The audit log was cleared'; Category = 'Tampering'; Subcategory = 'Audit Other Events'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Clearing the security log is a common anti-forensic technique. Investigate who cleared it and why.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-1102', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1102') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4616; Name = 'SystemTimeChanged'; DisplayName = 'The system time was changed'; Category = 'Tampering'; Subcategory = 'Audit Security State Change'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Time changes can disrupt log timeline correlation. Investigate intentional vs NTP sync.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4616', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4616') }
      )

      RegistryAuditing = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4657; Name = 'RegistryValueModified'; DisplayName = 'A registry value was modified'; Category = 'Registry Auditing'; Subcategory = 'Audit Registry'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Requires SACL-configured registry auditing. Track key/value paths for persistence modifications.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4657', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4657') }
      )

      ObjectAccess = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 4663; Name = 'ObjectAccessAttempt'; DisplayName = 'An attempt was made to access an object'; Category = 'Object Access'; Subcategory = 'Audit File System'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Requires SACL-configured auditing. Track file/folder access to sensitive paths (e.g. SAM, NTDS.dit).'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-4663', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4663') }
      )

      ShareAccess = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 5140; Name = 'ShareAccessed'; DisplayName = 'A network share object was accessed'; Category = 'Share Access'; Subcategory = 'Audit File Share'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Track remote share access. Correlate with logon events for the same source IP.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-5140', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5140') }
      )

      ShareChange = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 5142; Name = 'ShareAdded'; DisplayName = 'A network share object was added'; Category = 'Share Change'; Subcategory = 'Audit File Share'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('New network shares may indicate data staging or exfiltration preparation.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-5142', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5142') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 5143; Name = 'ShareModified'; DisplayName = 'A network share object was modified'; Category = 'Share Change'; Subcategory = 'Audit File Share'; Result = 'Success'; UtilityGroup = 'Logon'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-5143', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5143') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 5144; Name = 'ShareDeleted'; DisplayName = 'A network share object was deleted'; Category = 'Share Change'; Subcategory = 'Audit File Share'; Result = 'Success'; UtilityGroup = 'Logon'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-5144', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5144') }
      )

      NetworkFiltering = @(
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 5156; Name = 'WfpAllowedConnection'; DisplayName = 'The Windows Filtering Platform has allowed a connection'; Category = 'Network Filtering'; Subcategory = 'Audit Filtering Platform Connection'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Track allowed connections by application and destination.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-5156', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5156') }
        @{ LogName = 'Security'; ProviderName = 'Microsoft-Windows-Security-Auditing'; Id = 5157; Name = 'WfpBlockedConnection'; DisplayName = 'The Windows Filtering Platform has blocked a connection'; Category = 'Network Filtering'; Subcategory = 'Audit Filtering Platform Connection'; Result = 'Success'; UtilityGroup = 'Logon'; Notes = @('Blocked connections may indicate blocked malware C2 or data exfiltration attempts.'); References = @('https://learn.microsoft.com/en-us/previous-versions/windows/it-pro/windows-10/security/threat-protection/auditing/event-5157', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5157') }
      )
    }

    System = @{
      BootShutdown = @(
        @{
          LogName = 'System'
          ProviderName = 'Microsoft-Windows-Kernel-Power'
          Id = 41
          Name = 'UnexpectedShutdown'
          DisplayName = 'The system has rebooted without cleanly shutting down first'
          Category = 'Boot/Shutdown'
          Subcategory = 'Kernel-Power'
          Result = 'Success'
          UtilityGroup = 'BootShutdown'
          Notes = @('Unclean shutdown or system crash (BSOD). Check BugcheckCode / BugcheckParameter fields. Correlate with Application Error 1001 for more detail.')
          References = @(
            'https://learn.microsoft.com/en-us/windows/client-management/troubleshoot-event-id-41-restart',
            'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=41'
          )
        }
        @{
          LogName = 'System'
          ProviderName = 'BugCheck'
          Id = 1001
          Name = 'BugCheck'
          DisplayName = 'The computer has rebooted from a bugcheck'
          Category = 'Boot/Shutdown'
          Subcategory = 'BugCheck'
          Result = 'Success'
          UtilityGroup = 'BootShutdown'
          Notes = @('BSOD details including bugcheck code and parameters. Correlate with Kernel-Power 41.')
          References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1001')
        }
        @{
          LogName = 'System'
          ProviderName = 'User32'
          Id = 1074
          Name = 'SystemShutdown'
          DisplayName = 'The system has been shutdown or restarted by a user or process'
          Category = 'Boot/Shutdown'
          Subcategory = 'Shutdown'
          Result = 'Success'
          UtilityGroup = 'BootShutdown'
          Notes = @('Identifies the process and user that initiated a shutdown or restart.')
          References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1074')
        }
        @{
          LogName = 'System'
          ProviderName = 'EventLog'
          Id = 6005
          Name = 'EventLogStarted'
          DisplayName = 'The Event log service was started'
          Category = 'Boot/Shutdown'
          Subcategory = 'EventLog'
          Result = 'Success'
          UtilityGroup = 'BootShutdown'
          Notes = @('Indicates system startup. Pair with 6006 to determine uptime.')
          References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=6005')
        }
        @{
          LogName = 'System'
          ProviderName = 'EventLog'
          Id = 6006
          Name = 'EventLogStopped'
          DisplayName = 'The Event log service was stopped'
          Category = 'Boot/Shutdown'
          Subcategory = 'EventLog'
          Result = 'Success'
          UtilityGroup = 'BootShutdown'
          Notes = @('Indicates system shutdown. Pair with 6005 to determine uptime.')
          References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=6006')
        }
        @{
          LogName = 'System'
          ProviderName = 'EventLog'
          Id = 6008
          Name = 'UnexpectedShutdownPrevious'
          DisplayName = 'The previous system shutdown was unexpected'
          Category = 'Boot/Shutdown'
          Subcategory = 'EventLog'
          Result = 'Success'
          UtilityGroup = 'BootShutdown'
          Notes = @('Logged on the next boot after an unclean shutdown. Indicates the previous shutdown was unexpected.')
          References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=6008')
        }
      )

      Service = @(
        @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = 7000; Name = 'ServiceStartFailed'; DisplayName = 'A service failed to start'; Category = 'Service'; Subcategory = 'Service Control Manager'; Result = 'Failure'; UtilityGroup = 'Service'; Notes = @('Service failure to start - could be resource issue, dependency issue, or tampering with service binary.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7000') }
        @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = 7001; Name = 'ServiceDependencyFailed'; DisplayName = 'A service dependency failed to start'; Category = 'Service'; Subcategory = 'Service Control Manager'; Result = 'Failure'; UtilityGroup = 'Service'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7001') }
        @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = 7031; Name = 'ServiceTerminatedUnexpectedly'; DisplayName = 'A service terminated unexpectedly'; Category = 'Service'; Subcategory = 'Service Control Manager'; Result = 'Failure'; UtilityGroup = 'Service'; Notes = @('Service crash - may indicate an issue with the service or a targeted attack crashing security services.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7031') }
        @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = 7034; Name = 'ServiceTerminatedUnexpectedly2'; DisplayName = 'A service terminated unexpectedly'; Category = 'Service'; Subcategory = 'Service Control Manager'; Result = 'Failure'; UtilityGroup = 'Service'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7034') }
        @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = 7036; Name = 'ServiceStateChanged'; DisplayName = 'A service entered the running/stopped state'; Category = 'Service'; Subcategory = 'Service Control Manager'; Result = 'Success'; UtilityGroup = 'Service'; Notes = @('Service start/stop state transitions.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7036') }
        @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = 7040; Name = 'ServiceStartTypeChanged'; DisplayName = 'The start type of a service was changed'; Category = 'Service'; Subcategory = 'Service Control Manager'; Result = 'Success'; UtilityGroup = 'Service'; Notes = @('Changing from auto to disabled may be an attempt to neutralize security tools.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7040') }
        @{
          LogName = 'System'
          ProviderName = 'Service Control Manager'
          Id = 7045
          Name = 'ServiceInstalled'
          DisplayName = 'A service was installed in the system'
          Category = 'Service'
          Subcategory = 'Service Control Manager'
          Result = 'Success'
          UtilityGroup = 'Service'
          Fields = @{
            ServiceName = @{ XmlName = 'ServiceName'; Type = 'String'; Description = 'Name of the installed service.' }
            ImagePath = @{ XmlName = 'ImagePath'; Type = 'String'; Description = 'Binary path of the service executable.' }
            ServiceType = @{ XmlName = 'ServiceType'; Type = 'String'; Description = 'Service type (user mode service, kernel driver).' }
            StartType = @{ XmlName = 'StartType'; Type = 'String'; Description = 'Start type (auto start, demand start, disabled).' }
            AccountName = @{ XmlName = 'AccountName'; Type = 'String'; Description = 'Account the service runs as.' }
          }
          Notes = @(
            'New service installation is a common persistence mechanism.',
            'Check ImagePath for suspicious paths (Temp, AppData, etc.) or unusual executable names.',
            'Correlate with Security log event 4697 when available.'
          )
          References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7045')
        }
      )
    }

    Application = @{
      ApplicationCrash = @(
        @{ LogName = 'Application'; ProviderName = 'Application Error'; Id = 1000; Name = 'ApplicationCrash'; DisplayName = 'Application crash / faulting application'; Category = 'Application Crash'; Subcategory = 'Application Error'; Result = '$undefined'; UtilityGroup = 'ApplicationCrash'; Notes = @('Faulting application and module details. Check ExceptionCode for crash reason (e.g. 0xC0000005 = access violation).'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1000') }
        @{ LogName = 'Application'; ProviderName = 'Windows Error Reporting'; Id = 1001; Name = 'WerBucketReport'; DisplayName = 'Windows Error Reporting crash bucket event'; Category = 'Application Crash'; Subcategory = 'Windows Error Reporting'; Result = '$undefined'; UtilityGroup = 'ApplicationCrash'; Notes = @('WER bucket analysis metadata. Correlate with Application Error 1000.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1001') }
        @{ LogName = 'Application'; ProviderName = '.NET Runtime'; Id = 1026; Name = 'DotNetUnhandledException'; DisplayName = 'Unhandled .NET runtime exception'; Category = 'Application Crash'; Subcategory = '.NET Runtime'; Result = 'Failure'; UtilityGroup = 'ApplicationCrash'; Notes = @('Unhandled .NET exception with exception type and message.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1026') }
      )

      Installer = @(
        @{ LogName = 'Application'; ProviderName = 'MsiInstaller'; Id = 11707; Name = 'ProductInstalledSuccessfully'; DisplayName = 'Product installed successfully'; Category = 'Installer'; Subcategory = 'MsiInstaller'; Result = 'Success'; UtilityGroup = 'Installer'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=11707') }
        @{ LogName = 'Application'; ProviderName = 'MsiInstaller'; Id = 11708; Name = 'ProductInstallationFailed'; DisplayName = 'Product installation failed'; Category = 'Installer'; Subcategory = 'MsiInstaller'; Result = 'Failure'; UtilityGroup = 'Installer'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=11708') }
        @{ LogName = 'Application'; ProviderName = 'MsiInstaller'; Id = 11724; Name = 'ProductRemovedSuccessfully'; DisplayName = 'Product removed successfully'; Category = 'Installer'; Subcategory = 'MsiInstaller'; Result = 'Success'; UtilityGroup = 'Installer'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=11724') }
      )
    }

    PowerShell = @{
      All = @(
        @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; ProviderName = 'Microsoft-Windows-PowerShell'; Id = 400; Name = 'EngineStateAvailable'; DisplayName = 'Engine state changed to available'; Category = 'PowerShell'; Subcategory = 'Engine Lifecycle'; Result = 'Success'; UtilityGroup = 'PowerShell'; Notes = @('PowerShell engine became available (host started).'); References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=400') }
        @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; ProviderName = 'Microsoft-Windows-PowerShell'; Id = 403; Name = 'EngineStateStopped'; DisplayName = 'Engine state changed to stopped'; Category = 'PowerShell'; Subcategory = 'Engine Lifecycle'; Result = 'Success'; UtilityGroup = 'PowerShell'; Notes = @('PowerShell engine stopped.'); References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=403') }
        @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; ProviderName = 'Microsoft-Windows-PowerShell'; Id = 600; Name = 'ProviderStarted'; DisplayName = 'Provider started'; Category = 'PowerShell'; Subcategory = 'Provider Lifecycle'; Result = 'Success'; UtilityGroup = 'PowerShell'; Notes = @('PowerShell provider (e.g. WSMan) started. Treat with care - may indicate malicious remote sessions.'); References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=600') }
        @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; ProviderName = 'Microsoft-Windows-PowerShell'; Id = 800; Name = 'PipelineExecuted'; DisplayName = 'Pipeline executed'; Category = 'PowerShell'; Subcategory = 'Pipeline Execution'; Result = 'Success'; UtilityGroup = 'PowerShell'; Notes = @('A pipeline execution completed. Includes host application, command line, and user context.'); References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=800') }
        @{
          LogName = 'Microsoft-Windows-PowerShell/Operational'
          ProviderName = 'Microsoft-Windows-PowerShell'
          Id = 4103
          Name = 'ModuleLogging'
          DisplayName = 'Module logging'
          Category = 'PowerShell'
          Subcategory = 'Module Logging'
          Result = 'Success'
          UtilityGroup = 'PowerShell'
          Notes = @('Records pipeline execution details including parameters and values. Requires PowerShell Module Logging to be enabled in policy.')
          References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4103')
        }
        @{
          LogName = 'Microsoft-Windows-PowerShell/Operational'
          ProviderName = 'Microsoft-Windows-PowerShell'
          Id = 4104
          Name = 'ScriptBlockLogging'
          DisplayName = 'Script block logging'
          Category = 'PowerShell'
          Subcategory = 'Script Block Logging'
          Result = 'Success'
          UtilityGroup = 'PowerShell'
          Fields = @{
            ScriptBlockText = @{ XmlName = 'ScriptBlockText'; Type = 'String'; Description = 'Captured script block content (may be truncated at warning level).' }
            ScriptBlockId = @{ XmlName = 'ScriptBlockId'; Type = 'String'; Description = 'GUID identifier for the script block.' }
            Path = @{ XmlName = 'Path'; Type = 'String'; Description = 'Path of the script file if from a file, otherwise empty.' }
          }
          Notes = @(
            'The highest-value PowerShell event for IR. Contains captured script block content.',
            'Requires Script Block Logging policy to be enabled for full detail.',
            'At "Warning" log level, script blocks that match suspicious patterns are auto-logged even without full logging enabled.',
            'Base64-encoded commands appear here decoded for analysis.'
          )
          References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4104')
        }
        @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; ProviderName = 'Microsoft-Windows-PowerShell'; Id = 4105; Name = 'ScriptBlockInvocationStarted'; DisplayName = 'Script block invocation started'; Category = 'PowerShell'; Subcategory = 'Script Block Logging'; Result = 'Success'; UtilityGroup = 'PowerShell'; Notes = @('Indicates a script block started executing. Correlate with 4104 and 4106.'); References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4105') }
        @{ LogName = 'Microsoft-Windows-PowerShell/Operational'; ProviderName = 'Microsoft-Windows-PowerShell'; Id = 4106; Name = 'ScriptBlockInvocationCompleted'; DisplayName = 'Script block invocation completed'; Category = 'PowerShell'; Subcategory = 'Script Block Logging'; Result = 'Success'; UtilityGroup = 'PowerShell'; Notes = @('Indicates a script block completed execution. Correlate with 4104 and 4105.'); References = @('https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_logging_windows', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4106') }
      )
    }

    ScheduledTask = @{
      All = @(
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 100; Name = 'TaskStarted'; DisplayName = 'Task started'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/aa385123(v=vs.85)', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=100') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 101; Name = 'TaskStartFailed'; DisplayName = 'Task start failed'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Failure'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=101') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 102; Name = 'TaskCompleted'; DisplayName = 'Task completed'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=102') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 103; Name = 'ActionStartFailed'; DisplayName = 'Action start failed'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Failure'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=103') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 106; Name = 'TaskRegistered'; DisplayName = 'Task registered'; Category = 'TaskScheduler'; Subcategory = 'Task Registration'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; Notes = @('New scheduled task created. Check the task content for suspicious commands.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=106') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 107; Name = 'TaskTriggeredOnSchedule'; DisplayName = 'Task triggered by scheduler'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=107') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 108; Name = 'TaskTriggeredByEvent'; DisplayName = 'Task triggered by event'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; Notes = @('Event-triggered tasks are a common persistence method.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=108') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 110; Name = 'TaskTriggeredByUser'; DisplayName = 'Task triggered by user'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=110') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 111; Name = 'TaskTerminated'; DisplayName = 'Task terminated'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=111') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 118; Name = 'TaskTriggeredOnStartup'; DisplayName = 'Task triggered by computer startup'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=118') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 119; Name = 'TaskTriggeredOnLogon'; DisplayName = 'Task triggered by user logon'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=119') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 129; Name = 'TaskProcessCreated'; DisplayName = 'Task process created'; Category = 'TaskScheduler'; Subcategory = 'Task Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; Notes = @('Task created a new process. Contains the task name and launched process path.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=129') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 140; Name = 'TaskUpdated'; DisplayName = 'Task registration updated'; Category = 'TaskScheduler'; Subcategory = 'Task Registration'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; Notes = @('Task was modified. Check for changes to the action (command) or trigger.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=140') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 141; Name = 'TaskDeleted'; DisplayName = 'Task registration deleted'; Category = 'TaskScheduler'; Subcategory = 'Task Registration'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=141') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 142; Name = 'TaskDisabled'; DisplayName = 'Task disabled'; Category = 'TaskScheduler'; Subcategory = 'Task Registration'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; Notes = @('Task was disabled. May indicate an attacker disabling security tools or a defender disabling malicious persistence.'); References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=142') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 200; Name = 'ActionStarted'; DisplayName = 'Action started'; Category = 'TaskScheduler'; Subcategory = 'Action Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=200') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 201; Name = 'ActionCompleted'; DisplayName = 'Action completed'; Category = 'TaskScheduler'; Subcategory = 'Action Execution'; Result = 'Success'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=201') }
        @{ LogName = 'Microsoft-Windows-TaskScheduler/Operational'; ProviderName = 'Microsoft-Windows-TaskScheduler'; Id = 203; Name = 'ActionFailedToStart'; DisplayName = 'Action failed to start'; Category = 'TaskScheduler'; Subcategory = 'Action Execution'; Result = 'Failure'; UtilityGroup = 'ScheduledTask'; References = @('https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=203') }
      )
    }

    Sysmon = @{
      All = @(
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 1; Name = 'ProcessCreation'; DisplayName = 'Process creation'; Category = 'Sysmon'; Subcategory = 'Process'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Process creation with full command line, hashes, parent process, user, and integrity level.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=1') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 2; Name = 'FileCreationTimeChanged'; DisplayName = 'File creation time changed'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('File creation time was changed. May indicate timestomping (anti-forensic technique).'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=2') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 3; Name = 'NetworkConnection'; DisplayName = 'Network connection'; Category = 'Sysmon'; Subcategory = 'Network'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('TCP/UDP network connection with source/destination IP, port, and owning process.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=3') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 4; Name = 'SysmonServiceStateChanged'; DisplayName = 'Sysmon service state changed'; Category = 'Sysmon'; Subcategory = 'Service'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Sysmon service started or stopped. Stopping Sysmon may indicate tampering.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=4') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 5; Name = 'ProcessTerminated'; DisplayName = 'Process terminated'; Category = 'Sysmon'; Subcategory = 'Process'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Process termination. Correlate PID with event 1.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=5') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 6; Name = 'DriverLoaded'; DisplayName = 'Driver loaded'; Category = 'Sysmon'; Subcategory = 'Driver'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Kernel driver loaded. Signed but uncommon drivers may indicate rootkits or vulnerable drivers being exploited.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=6') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 7; Name = 'ImageLoaded'; DisplayName = 'Image loaded'; Category = 'Sysmon'; Subcategory = 'Image'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('DLL/module loaded into a process. High volume event; configure filtering to focus on suspicious modules.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=7') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 8; Name = 'CreateRemoteThread'; DisplayName = 'CreateRemoteThread'; Category = 'Sysmon'; Subcategory = 'Thread'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Remote thread creation. Classic process injection indicator.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=8') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 9; Name = 'RawAccessRead'; DisplayName = 'RawAccessRead'; Category = 'Sysmon'; Subcategory = 'Disk'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Raw disk access. Credential dumping tools often read disk sectors directly.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=9') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 10; Name = 'ProcessAccess'; DisplayName = 'ProcessAccess'; Category = 'Sysmon'; Subcategory = 'Process'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Process access (OpenProcess). LSASS access with PROCESS_VM_READ rights indicates credential dumping.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=10') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 11; Name = 'FileCreate'; DisplayName = 'FileCreate'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('File creation. Monitor for executable/.dll file drops in Temp, AppData, or Startup folders.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=11') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 12; Name = 'RegistryCreateDelete'; DisplayName = 'Registry object created or deleted'; Category = 'Sysmon'; Subcategory = 'Registry'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Registry key/value creation or deletion. Monitor common persistence locations (Run keys, services, etc.).'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=12') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 13; Name = 'RegistryValueSet'; DisplayName = 'Registry value set'; Category = 'Sysmon'; Subcategory = 'Registry'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Registry value modified. Track modifications to common persistence and configuration keys.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=13') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 14; Name = 'RegistryRename'; DisplayName = 'Registry object renamed'; Category = 'Sysmon'; Subcategory = 'Registry'; Result = 'Success'; UtilityGroup = 'Sysmon'; References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=14') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 15; Name = 'FileCreateStreamHash'; DisplayName = 'FileCreateStreamHash'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Alternate data stream (ADS) creation. Mark of the Web / Zone.Identifier is common. Other streams may indicate data hiding.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=15') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 16; Name = 'SysmonConfigChanged'; DisplayName = 'Sysmon configuration changed'; Category = 'Sysmon'; Subcategory = 'Configuration'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Sysmon configuration was updated. Unexpected changes may indicate tampering.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=16') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 17; Name = 'PipeCreated'; DisplayName = 'Pipe created'; Category = 'Sysmon'; Subcategory = 'Pipe'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Named pipe creation. Cobalt Strike and other C2 frameworks use named pipes for inter-process communication.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=17') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 18; Name = 'PipeConnected'; DisplayName = 'Pipe connected'; Category = 'Sysmon'; Subcategory = 'Pipe'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Named pipe connection. Correlate with event 17.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=18') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 19; Name = 'WmiEventFilterActivity'; DisplayName = 'WMI event filter activity'; Category = 'Sysmon'; Subcategory = 'WMI'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('WMI event filter created/modified/deleted. Common persistence mechanism (WMI subscription).'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=19') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 20; Name = 'WmiEventConsumerActivity'; DisplayName = 'WMI event consumer activity'; Category = 'Sysmon'; Subcategory = 'WMI'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('WMI event consumer created/modified/deleted. Correlate with event 19 for persistence detection.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=20') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 21; Name = 'WmiFilterToConsumerBinding'; DisplayName = 'WMI consumer to filter binding'; Category = 'Sysmon'; Subcategory = 'WMI'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('WMI filter-to-consumer binding created/deleted. Correlate with events 19 and 20.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=21') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 22; Name = 'DnsQuery'; DisplayName = 'DNS query'; Category = 'Sysmon'; Subcategory = 'DNS'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('DNS query from a process. Monitor for queries to known-malicious, newly registered, or algorithmically generated domains.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=22') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 23; Name = 'FileDelete'; DisplayName = 'File delete'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('File deletion with archiving option (saves deleted file contents). Monitor for evidence destruction.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=23') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 24; Name = 'ClipboardChanged'; DisplayName = 'Clipboard changed'; Category = 'Sysmon'; Subcategory = 'Clipboard'; Result = 'Success'; UtilityGroup = 'Sysmon'; References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=24') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 25; Name = 'ProcessTampering'; DisplayName = 'Process tampering'; Category = 'Sysmon'; Subcategory = 'Process'; Result = 'Success'; UtilityGroup = 'Sysmon'; Notes = @('Process tampering via process hollowing or herpaderping techniques detected.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=25') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 26; Name = 'FileDeleteDetected'; DisplayName = 'File delete detected'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=26') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 27; Name = 'FileBlockExecutable'; DisplayName = 'File block executable'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=27') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 28; Name = 'FileBlockShredding'; DisplayName = 'File block shredding'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=28') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 29; Name = 'FileExecutableDetected'; DisplayName = 'File executable detected'; Category = 'Sysmon'; Subcategory = 'File'; Result = 'Success'; UtilityGroup = 'Sysmon'; References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=29') }
        @{ LogName = 'Microsoft-Windows-Sysmon/Operational'; ProviderName = 'Microsoft-Windows-Sysmon'; Id = 255; Name = 'SysmonError'; DisplayName = 'Sysmon error'; Category = 'Sysmon'; Subcategory = 'Error'; Result = 'Failure'; UtilityGroup = 'Sysmon'; Notes = @('Sysmon encountered an internal error. May indicate resource pressure or tampering.'); References = @('https://learn.microsoft.com/en-us/sysinternals/downloads/sysmon', 'https://www.ultimatewindowssecurity.com/securitylog/encyclopedia/event.aspx?eventID=255') }
      )
    }
  }
}
