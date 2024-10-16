function runFile() {
  wshShell = new ActiveXObject('WScript.Shell');
  wshShell.run('c:/windows/system32/notepad.exe', 1, false);
}
