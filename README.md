# PMPHyperVClone

PowerShell script to clone Hyper-V VMs from one host to another.

## Requirements

- SYSTEM account of source host must be local admin on target host: VMMS service runs as SYSTEM by default and that's used for Export-VM. 
    - This could be negated by exporting locally first, then copy to target host
- Invoking user be in the Hyper-V Administrators group on the source host or execute the script as administrator / elevated
