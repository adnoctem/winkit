## 1.0.0 (2026-08-10)

### Features

* add `host` and `path` library ([3b1301a](https://github.com/adnoctem/winkit/commit/3b1301aeb465746684531fa9144018590d30bee6))
* add `Repair-WSL` and start Office scripts ([d481bf4](https://github.com/adnoctem/winkit/commit/d481bf437d1ef6d3e20d69e6cf9202f2d348debf))
* add `Update-AutoDNSZones` ([6b8b2cb](https://github.com/adnoctem/winkit/commit/6b8b2cb0133e47ae342a9c08263c6f8edbca03e0))
* add expand configuration in `Update-AutoDNSZones` ([10c0893](https://github.com/adnoctem/winkit/commit/10c0893d5ad08af5980e9b1531e942918e4f07cd))
* add hosts blocklist, secure DNS, TCP, power plan, and disk optimization scripts ([f434c32](https://github.com/adnoctem/winkit/commit/f434c32ec4f460ef0c8cdeb62c5c09e6be142a2a))
* add IPv6 features to `networking.ps1` ([220ee62](https://github.com/adnoctem/winkit/commit/220ee62ce122bfc6bdcbb1d366ac8a66eccd8d29))
* add majority of planned features in `scripts` ([91be888](https://github.com/adnoctem/winkit/commit/91be888606be1107ee156234da2e5dae6acdfe58))
* add script to change text file encoding ([af3aaf3](https://github.com/adnoctem/winkit/commit/af3aaf3c246d63f57ab49f7ab558379f86ccdf64))
* finish repository housekeeping and add plethora of features ([6fc132c](https://github.com/adnoctem/winkit/commit/6fc132cabbfcf30e71d68cc85bbbbfc348f63f4f))
* harden and standardize all scripts ([f46ae88](https://github.com/adnoctem/winkit/commit/f46ae88a7105a39dc32f7b560f055b60d2213ad1))
* rewrite `networking.ps1` ([630c76c](https://github.com/adnoctem/winkit/commit/630c76c6c7cf606de99a80cac20547c32f71efc7))
* **scripts/Windows:** add `Find-OffHoursActivity` ([a66ce65](https://github.com/adnoctem/winkit/commit/a66ce65d662a9fabae3525049f063d0ebaf5abe4))
* **scripts:** add `Export-AD` script ([e03613f](https://github.com/adnoctem/winkit/commit/e03613f40b51413bebc2739428c0faba06833983))
* **scripts:** add Find-MACAddress and migrate console output to Write-Log ([d172c33](https://github.com/adnoctem/winkit/commit/d172c332bf058c99fcf3c7f69cdcadcce3b07686))
* **scripts:** add script `New-WGConfigurations` ([a1d6425](https://github.com/adnoctem/winkit/commit/a1d64256e1d1f86ef7b65bfd3f74975cac0bf199))
* **security:** add semantic Windows Event Log API and integration ([09138be](https://github.com/adnoctem/winkit/commit/09138bea3a5cb69c2d7c8b830445977304c0a79c))

### Bug Fixes

* **ci:** grant dispatch job contents:write for repository_dispatch ([505c266](https://github.com/adnoctem/winkit/commit/505c26610a2378e4fb7671635429be95bc4e23e9))
* correct DHCP import Cmdlet ([1a0c404](https://github.com/adnoctem/winkit/commit/1a0c4043085b42201d1bcdb4322a3fb3e65c0641))
* **lib/settings:** add missing `function` keyword ([9501a49](https://github.com/adnoctem/winkit/commit/9501a498438c2a003a4f218b93dc72a61956f03e))
* repair module import paths and registry value existence checks ([ba1bd50](https://github.com/adnoctem/winkit/commit/ba1bd502f965b91a7e49b10ccd52c6ae21b1d98a))
* **scripts:** fix `Export-AD` with PS2 compatibility ([d1bca4a](https://github.com/adnoctem/winkit/commit/d1bca4ae84426ebc50cdc0bde20949e986589248))
* **scripts:** prevent false-positive DNS change detection and clean up output ([393253d](https://github.com/adnoctem/winkit/commit/393253d6dddd1a1f969483fe5b72de62c1cf6c47))
