Public Function UStax(income)
    Dim State
    State = 0.0122
    
    Select Case income
    Case Is >= 523601
        UStax = 153687 + 0.37 * (income - 523600)
    Case Is >= 209425
        UStax = 43718 + 0.35 * (income - 209424)
    Case Is >= 164924
        UStax = 29478 + 0.32 * (income - 164923)
    Case Is >= 86375
        UStax = 10627 + 0.24 * (income - 86374)
    Case Is >= 40525
        UStax = 540 + 0.22 * (income - 40524)
    Case Is >= 9950
        UStax = 100 + 0.12 * (income - 9949)
    Case Else
        UStax = 0.1 * income
    End Select
End Function

Public Function UKtax(income)
    Select Case income
    Case Is >= 202815
        UKtax = 68557 + 0.45 * (income - 202815)
    Case Is >= 67970
        UKtax = 7540 + 0.4 * (income - 67970)
    Case Is >= 16995
        UKtax = 0.2 * (income - 16994)
    Case Else
        UKtax = 0
    End Select
End Function
