Attribute VB_Name = "basUsers"
Option Explicit

Public Enum wis_Permissions

'    perNoPermissions = 0
'    perFullPermissions = 1
'
'    perSBCACreate = 2    '2,147,483,647 long 1,073,741,824
'    perSBCATrans = 4
'    perSBCAUnDo = 8
'    perSBCAView = 16
'
'    perDepCreate = 32
'    perDepTrans = 64
'    perDepUnDo = 128
'    perDepView = 256
'    '2,147,483,647 long 1,073,741,824
'
'    perPDCreate = 512   '256 * 2
'    perPDTrans = 1024   '256 * 4
'    perPDUnDo = 2048    '256 * 8
'    perPDView = 4096    '256 * 16
'
'    perMemCreate = 8192 '4096 * 2
'    perMemTrans = 16384 '4096 * 4
'    perMemUnDo = 32768  '4096 * 8
'    perMemView = 65536  '4096 * 16
'
'    perLoanCreate = 131072 '65536 * 2
'    perLoanTrans = 262144   '65536 * 4
'    perLoanUnDo = 524288    '65536 * 8
'    perLoanView = 1048576   '65536 * 16
'
'    perUserCreate = 2097152 '1048576 * 2
'    perUserDelete = 4194304 '1048576 * 4
'
'    perTransPaybles = 8388608 '4194304 * 2
'    perUndoPaybles = 16777216 '4194304 *4
'
'    perBankAccCreate = 33554432 '16777216  * 2
'    perBankAccTrans = 67108864  '16777216  * 4
'
'    perReport = 134217728 ' 67108864 *2
'
'    perPigmyOperator = 134217728 * 2
'    'perPigmyOperator = perPDCreate Or perPDTrans Or perPDUnDo + perPDView
    
    perOnlyWaves = 1024
    perBankAdmin = 256
    perCreateLedger = 128
    perModifyAccount = 64
    perCreateAccount = 32
    perPassingOfficer = 16
    perCashier = 8
    perClerk = 4
    perPigmyAgent = 2
    perReadOnly = 1
    
    perNoPerms = 0

End Enum

Public Enum wis_Operations
    wisUserCreation = 1
    wisTransaction = 2
    wisUndoTransaction = 4
    wisReadTransact = 8
    wisPigmy = 16
End Enum
