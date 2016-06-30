Attribute VB_Name = "M90_DAF_Various"



''presumably assigned to one button (or some event you specify) to disable double-clicking:
 Sub DoubleClickDisable()
 Application.OnDoubleClick = False ''"TurnOff"
 End Sub

'presumably assigned to another button (or some event you specify) to re-enable double-clicking:
 Sub DoubleClickEnable()
 Application.OnDoubleClick = True
 End Sub

