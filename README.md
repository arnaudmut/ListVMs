# ListVms

Retrieves a list of virtual machines from hostIP hosts or vCenter Server, extracting a subset of vm properties values returned by Get-View.
Hosts names(IPs) and credentials are saved in a XMl file ( config.xml for this script). 
The script saves reports to disk and automatically displays the reports converted to HTML by invoking the default browser.
Reports are exported to Excel using Import-Excel Modules from https://github.com/dfinke/ImportExcel.
You can install this Module from PowerShell Gallery : https://www.powershellgallery.com/packages/ImportExcel/5.3.4 

## Usage

```Powershell
.\listvms <sortBy>:[Optional] <exportTo>:[optional]
```

__Change params to suit your needs__   

### Prerequisites
* VMware Infrastructure
* ImportExcel Modules



## Contributing
1. Fork it!
2. Create your feature branch: `git checkout -b my-new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin my-new-feature`
5. Submit a pull request :D

## History
TODO: Write history

## Credits
TODO: Write credits

## License
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details
## Authors

* **Arnaud Mutana**  - [kardudu](https://www.arnaudmut.fr "Welcome") &copy; 
