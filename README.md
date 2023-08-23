# ADCSmaintenance
Windows ADCS database maintenance script. This script manages the ADCS database. It will remove expired certificates, 
remove failed certificates and remove pending requests. 
Before a certificates will be deleted, the script has the option to export those issued certificates. 
The configuration file will be created with the config-camaintenance.ps1 script

The script can run as a schedule task in the system context
