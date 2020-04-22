function seeDrive() {
  const clientFolder = DriveApp.getFolderById('1Dr02Xk4U9H_QPNPbD-zXMoaZh-40ZJVq')
  const copyAttyFolder = DriveApp.getFolderById('1CsSEPi5X4adcBbFNew0MUOEiM0VgYqGr')
  const allClientFolders = clientFolder.getFolders()
  const allCopyAttyFolders = clientFolder.getFolders()
  const folderCheck = clientFolder.getFoldersByName('test')
}
