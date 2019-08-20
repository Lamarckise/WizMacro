var WizMacro = (function () {
  var mainframe = new ActiveXObject("BZWhll.WhllObj");
  var session = mainframe.ActiveSession;
  var screen = session.Screen;
  
  return {
    send: function(text) {
      screen.SendKeys(text);
      mainframe.WaitReady(10, 1);
    },
    read: function(fromX, fromY, toX, toY) {
      return screen.Area(fromX, fromY, toX, toY);
    }
  }
})();
