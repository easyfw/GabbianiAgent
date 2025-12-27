object GabbianiAgent: TGabbianiAgent
  OldCreateOrder = False
  DisplayName = 'Gabbiani Data Agent'
  Left = 298
  Top = 175
  Height = 150
  Width = 215
  object MyComm: TVaComm
    FlowControl.OutCtsFlow = False
    FlowControl.OutDsrFlow = False
    FlowControl.ControlDtr = dtrDisabled
    FlowControl.ControlRts = rtsDisabled
    FlowControl.XonXoffOut = False
    FlowControl.XonXoffIn = False
    FlowControl.DsrSensitivity = False
    FlowControl.TxContinueOnXoff = False
    DeviceName = 'COM%d'
    OnRxChar = MyCommRxChar
    Version = '1.3.5.0'
    Left = 56
    Top = 32
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 144
    Top = 40
  end
end
