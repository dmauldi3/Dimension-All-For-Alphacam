using AlphaCAMMill;
using System;


namespace AlphacamAddinSample.Models
{
  public sealed class AlphacamEvents : IDisposable
  {
    private bool _bDisposed;
    private MemoryManager _memMgr;
    private IAlphaCamApp _comApp;
    private AddInInterface _comInterface;
    private Addin _objAddin;
    private CommandItem _comCommand;

    public AlphacamEvents(IAlphaCamApp app)
    {
      this._memMgr = new MemoryManager();
      this._comApp = app;
      using (MemoryManager memoryManager = new MemoryManager())
      {
        _comInterface = _memMgr.Add(memoryManager.Add(app.Frame).CreateAddInInterface());
        
        _comInterface.InitAlphacamAddIn += OnInitAddin;
        _comInterface.BeforeClose += OnBeforeClose;
      }
    }

    private void OnBeforeClose(EventData data)
    {
      this._objAddin.ResetRunCount();
      data.ReturnCode = 0;
    }

    private void OnInitAddin(AcamInitAddInAction action, EventData data)
    {
      using (MemoryManager memoryManager = new MemoryManager())
      {
       _objAddin = new Addin(_comApp.Application);
        IFrame frame =  memoryManager.Add(_comApp.Frame);
        _comCommand = _memMgr.Add(frame.CreateCommandItem());
        
        _comCommand.OnCommand += (OnCommand);
        _comCommand.OnUpdate += (OnUpdate);
        frame.AddMenuItem43("DIMENSION ALL", "CMD_MYCOMMAND", AcamCommand.acamCmdUTILS_ADDINS, true, "My Addins", frame.LastMenuCommandID, this._comCommand);
      }
      data.ReturnCode = 0;
    }

    private AcamOnUpdateReturn OnUpdate() => AcamOnUpdateReturn.acamOnUpdate_UncheckedEnabled;

    private void OnCommand() => _objAddin.DimensionAll();

    public void Dispose() => DisposeClass();

    ~AlphacamEvents() => DisposeClass();

    private void DisposeClass()
    {
      if (_bDisposed)
        return;
      _memMgr.Dispose();
      _bDisposed = true;
      GC.SuppressFinalize( this);
    }
  }
}