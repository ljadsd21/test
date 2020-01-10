(function(){
  var WSHShell = new ActiveXObject("WScript.Shell");
  var Env = WSHShell.Environment("PROCESS");
  var s = "";
  var colVars = new Enumerator(Env);
  for(; !colVars.atEnd(); colVars.moveNext()) {
    s += colVars.item() + "\n";
  }
WScript.Echo(s);
})();