# energy-tools
Energy Tools for Energy Star and ErP Lot 3

This program is designed to collect the system profile and calculate the results of Energy Star (5.2 & 6.0 & 7.0 & 8.0) and ErP Lot 3 (Jan. 2016).

## Usage

```
$ energy-tools -h
usage: energy-tools [-h] [-d] [-e] [-r] [-p PROFILE] [-t TEST]

Energy Tools 1.7 for Energy Star 5/6/7/8 and ErP Lot 3

optional arguments:
  -h, --help            show this help message and exit
  -d, --debug           print debug messages
  -e, --excel           generate Excel file
  -r, --report          generate report file
  -s, --simulate        simulate 4G ram (Not support in Snap package.)
  -p PROFILE, --profile PROFILE
                        specify profile
  -t TEST, --test TEST  use test case
```

## Snap Package

[![Snap Status](https://build.snapcraft.io/badge/fourdollars/energy-tools.svg)](https://build.snapcraft.io/user/fourdollars/energy-tools)

[![Get it from the Snap Store](https://snapcraft.io/static/images/badges/en/snap-store-white.svg)](https://snapcraft.io/energy-tools)

```
$ snap install energy-tools
$ snap connect energy-tools:hardware-observe
$ snap connect energy-tools:home
$ snap connect energy-tools:mount-observe
```

* Connecting ':hardware-observe' interface allows it to get the hardware information.
* Connecting ':home' interface allows it to read/write files under home directory.
* Connecting ':mount-observe' interface allows it to check the system disk.

## Ubuntu PPA

```
$ sudo add-apt-repository ppa:fourdollars/energy-tools
$ sudo apt-get update
$ sudo apt-get install energy-tools
```

## WARNING

This tool just works for general computers right now.
