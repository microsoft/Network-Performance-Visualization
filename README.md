# Description
Network-Data-Visualization is a tool used to visualize the data produced by various network performance analysis tools. Currently, the project supports the tools NTTTCP, LATTE, and CTStraffic.

Though these tools are quite capable, reading and interpreting their raw output files is a tedious and time consumimg task. Additionally, network performance tests are often run for multiple iterations in order to mitigate the effects of random variance, and this can create directories full of these dense data files.   

Given a directory full of NTTTCP, LATTE, or CTStraffic data files, this tool will parse the raw data, analyze it, and then create tables and charts in excel which provide useful visualizations of that data. 

<p align="center">
  <img src="/images/latency-histogram.PNG" title="Latency Histogram" width=75% height=75%>
  <img src="/images/throughput-quartiles.PNG" title="Throughput Quartiles" width=75% height=75%>
  <img src="/images/latency-percentiles.PNG" title="Latency Percentiles" width=75% height=75%>
</p>

The tool can aggregate data from multiple iterations of network performance monitoring tools and it can be given two directories in order to create side by side comparisons of performance measures before and after system changes. 



# Installation
## Manual Installation
Download this repo to your machine, and then move the Network-Performance-Visualization folder to C:\Program Files\WindowsPowerShell\Modules. After moving the folder, PowerShell should automatically install the module. You can double check that everything was installed correctly by running the command
```PowerShell
Get-Module -ListAvailable
```
and checking that `Network-Performance-Visualization` is listed among the available modules.
# Usage
The `Network-Performance-Visualization` module exports one command called `New-Visualization`.  
For help and options when running this command directly, use:
```PowerShell
Get-Help New-Visualization
```
# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
