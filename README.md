### MSI Plugin for Total Commander
This plugin can be installed into Total Commander. It allows to view MSI archives as if they were a ZIP file.

Because MSI files are databases, the plugin turns them into virtual files.
 * If a database table contains rows with a stream, it is shown as a folder (named after the table) and each row is a single file in that folder.
 * Otherwise, the database table is shown as a virtual UTF8-encoded CSV file.

### Build Requirements
To build the MSI plugin, you need to have one of these build environments
* Visual Studio 202x
* WDK 6001
Also, the following tool is needed to be in your PATH:
* zip.exe (https://sourceforge.net/projects/infozip/files/)

1) Make a new directory, e.g. C:\Projects
```
md C:\Projects
cd C:\Projects
```

2) Clone the common library
```
git clone https://github.com/ladislav-zezula/Aaa.git
```

3) Clone and build the WCX_MSI plugin
```
git clone https://github.com/ladislav-zezula/wcx_msi.git
cd wcx_msi
make-msvc.bat /web
```
The installation package "wcx_msi.zip" will be in the project directory.

4) Install the plugin.
 * Locate the wcx_msi.zip file in Total Commander
 * Double-click on it with the mouse (or press Ctrl+PageDown)
 * Total Commander will tell you that the archive contains a MSI packer plugin
   and asks whether you want to install it. Click on "Yes".
 * Next, Total Commander asks where do you want to install the plugin.
   Just confirm by clicking "OK"
 * Total commander will then install the plugin
 * After the installation, a configuration dialog opens. By default,
   the "msi" extension should be selected and the proper plugin should be focused.
   Just click "OK"
 * The plugin should now be fully operational. Try it by locating a MSI file
   and double-clicking it in Total Commander
 * Alternatively, you can press Ctrl+PageDown on a MSI file (regardless of its extension)
