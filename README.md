# energy-tools
Energy Tools for Energy Star and ErP Lot 3

This program is designed to collect the system profile and calculate the results of Energy Star (5.2 & 6.0 & 7.0 & 8.0) and ErP Lot 3 (Jan. 2016).

## Usage
```
$ energy-tools -h
usage: energy-tools [-h] [-d] [-e] [-r] [-p PROFILE] [-t TEST]

Energy Tools 1.6 for Energy Star 5/6/7/8 and ErP Lot 3

optional arguments:
  -h, --help            show this help message and exit
  -d, --debug           print debug messages
  -e, --excel           generate Excel file
  -r, --report          generate report file
  -s, --simulate        simulate 4G ram
  -p PROFILE, --profile PROFILE
                        specify profile
  -t TEST, --test TEST  use test case
```

You need to execute this command with root permission so it can collect the hardware information.

`$ sudo energy-tools`

If you are using some valid profile, you don't need to execute this command with root permission.

`$ energy-tools -p MySystem.profile`

## Snap Package
`$ snap install --beta --devmode energy-tools`

## Ubuntu PPA
```
$ sudo add-apt-repository ppa:fourdollars/energy-tools
$ sudo apt-get update
```

## WARNING

This tool just works for general computers right now.
