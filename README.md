# Windows-Time-Checker
A simple programme that correct local system time on Windows

The programme needs admin permission to make it start when Windows boot

put main.py and config.yaml in the same path

This programme also allows you to use your own application to check time,edit the app's path to the Path line in config.yaml and change Use_excternal_path line to "yes"

you can set whether you wants the progamme start automatically when windows boot by editing the line lauch_when_device_start in config.yaml,but make sure the programme has the admin permission otherwise auto run will not available.
