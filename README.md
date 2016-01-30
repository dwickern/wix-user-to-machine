wix-user-to-machine
===================

Upgrade a WiX installation from per-user to per-machine

Out of the box, MSI will not upgrade a per-user product to per-machine.
This is a [well](http://stackoverflow.com/questions/678002/how-do-i-fix-the-upgrade-logic-of-a-wix-setup-after-changing-installscope-to-pe)
[known](http://stackoverflow.com/questions/12048032/why-major-upgrade-does-not-upgrade-previous-per-machine-installation)
[problem](http://stackoverflow.com/questions/11119838/wix-installer-cant-upgrade-from-previously-installed-windows-installer-sw).
This example leverages burn to work around this limitation.

Projects:
- [UpgradeTestApplication](UpgradeTestApplication): A dummy application to install
- [PerUserSetup](PerUserSetup): WiX MSI with `perUser` scope
- [PerMachineSetup](PerMachineSetup): WiX MSI with `perMachine` scope
- [UninstallRelatedProducts](UninstallRelatedProducts): Command-line application which uninstalls products with a given upgrade code
- [PerMachineBootstrapper](PerMachineBootstrapper): WiX bootstrapper which combines `UninstallRelatedProducts` and `PerMachineSetup`

To reproduce the upgrade problem, install `PerUserSetup` then `PerMachineSetup`.
There will be two ARP entries since MSI did not perform an upgrade.

If you install `PerUserSetup` then `PerMachineBootstrapper`, the per-user installation will be properly removed.
There will only be one ARP entry for `PerMachineBootstrapper`.

## License

See [LICENSE](LICENSE.md) (MIT).
