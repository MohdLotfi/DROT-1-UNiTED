VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub HEMOS_V2()


Dim EnerSup() As Double  
Dim CustSup() As Double  
Dim EnerBlocks() As Double  
Dim ValueBlocks() As Double  
Dim NSlotBlocks() As Integer  
Dim OrdSup() As Integer  
Dim OrdSupFlag() As Boolean 
Dim EnerLoad() As Double  
Dim AlocateBlock() As Integer  
Dim AlocateEnerLoad() As Double  
Dim AlocateSup() As Double  
Dim BlockFlag() As Boolean 

n_slot = Worksheets("EnerSup").Cells(1, 3)  
n_sup = Worksheets("EnerSup").Cells(2, 3)   
n_blocks = Worksheets("EnerBlocks").Cells(1, 3)   

ReDim EnerSup(1 To n_slot, 1 To n_sup)
ReDim CustSup(1 To n_slot, 1 To n_sup)
ReDim ValueBlocks(1 To n_slot, 1 To n_blocks)
ReDim OrdSup(1 To n_slot, 1 To n_sup)
ReDim OrdSupFlag(1 To n_sup)
ReDim OrdSup(1 To n_slot, 1 To n_sup)
ReDim EnerLoad(1 To n_slot)
ReDim AlocateBlock(1 To n_slot, 1 To n_blocks)
ReDim AlocateEnerLoad(1 To n_slot, 1 To n_blocks)
ReDim AlocateSup(1 To n_slot, 1 To n_sup)
ReDim BlockFlag(1 To n_blocks)

ReDim EnerBlocks(1 To n_slot, 1 To n_blocks)
ReDim NSlotBlocks(1 To n_blocks)

nIterTotal = 10

Dim sumsGrid() As Double
Dim sumsPs() As Double
Dim sumsPo() As Double

ReDim sumsG(1 To nIterTotal)
ReDim sumsS(1 To nIterTotal)
ReDim sumsP(1 To nIterTotal)

For iterTotalN = 1 To nIterTotal

    'leitura de vetores
    'Matriz Enerform
    For i = 1 To n_slot
        For j = 1 To n_forn
            EnerSup(i, j) = Worksheets("EnerSup").Cells(6 + i, 2 + j)
        Next j
    Next i
    
    'Matriz CustSup
    For i = 1 To n_slot
        For j = 1 To n_forn
            CustSup(i, j) = Worksheets("CustSup").Cells(6 + i, 2 + j)
        Next j
    Next i
    
    'Matriz ValueBlocks
    For i = 1 To n_slot
        For j = 1 To n_blocks
            ValueBlocks(i, j) = Worksheets("ValueBlocks").Cells(6 + i, 2 + j)
        Next j
    Next i
    
    'Matriz EnerBlocks
    For i = 1 To n_blocks
        NSlotBlocks(i) = Worksheets("EnerBlocks").Cells(5, 2 + i)
    Next i
    
    For i = 1 To n_slot
        For j = 1 To n_blocks
            EnerBlocks(i, j) = Worksheets("EnerBlocks").Cells(6 + i, 2 + j)
        Next j
    Next i
    
    For i = 1 To n_slot
        For j = 1 To n_forn
            MinCustForn = 999999999
            For k = 1 To n_forn
                If Not (OrdSupFlag(k)) Then
                    If CustSup(i, k) < MinCustForn Then
                        OrdSup(i, j) = k
                        MinCustForn = CustSup(i, k)
                    End If
                End If
            Next k
            OrdSupFlag(OrdSup(i, j)) = True
        Next j
        For k = 1 To n_forn
            OrdSupFlag(k) = False
        Next k
    Next i
    
    
    For ii = 1 To n_blocks 
        ProveitoMax = 0
        Bloco = 0
        Slot = 0
        For i = 1 To n_blocks 
            If BlockFlag(i) = False Then 
                For j = 1 To n_slot 
                    Proveito = 0
                    For l = 1 To NSlotBlocks(i)
                        If (j + l - 1) > n_slot Then Exit For
                        EnerSupAcc = 0
                        Custo_e = 0
                        supply = EnerLoad(j + l - 1) + EnerBlocks(l, i) 
                        For k = 1 To n_forn
                            EnerSupAcc = EnerSupAcc + EnerSup(j + l - 1, OrdSup(j + l - 1, k))
                            If supply > EnerSupAcc Then
                                Custo_e = Custo_e + EnerSup(j + l - 1, OrdSup(j + l - 1, k)) * CustSup(j + l - 1, OrdSup(j + l - 1, k))
                            Else
                                FracEnerSup = supply - (EnerSupAcc - EnerSup(j + l - 1, OrdSup(j + l - 1, k)))
                                Custo_e = Custo_e + FracEnerSup * CustSup(j + l - 1, OrdSup(j + l - 1, k))
                                Custo_e_kW = Custo_e / supply
                                Exit For
                            End If
                        Next k
                        Proveito = Proveito + ValueBlocks(j + l - 1, i) - Custo_e_kW
                    Next l
                    
                    If Proveito > ProveitoMax Then
                        ProveitoMax = Proveito
                        Bloco = i
                        Slot = j
                    End If
                Next j
            End If
        Next i
        For l = 1 To NSlotBlocks(Bloco)
            AlocateBlock(1 + (Slot + l - 2) Mod 24, Bloco) = ii
            Worksheets("AlocateBlock").Cells(6 + 1 + (Slot + l - 2) Mod 24, 2 + Bloco) = 1 + (Slot + l - 2) Mod 24
            AlocateEnerLoad(1 + (Slot + l - 2) Mod 24, Bloco) = EnerBlocks(l, Bloco)
            EnerLoad(1 + (Slot + l - 2) Mod 24) = EnerLoad(1 + (Slot + l - 2) Mod 24) + EnerBlocks(l, Bloco)
            BlockFlag(Bloco) = True
        Next l
    Next ii
    
    For j = 1 To n_slot 
        supply = EnerLoad(j) 
        EnerSupAcc = 0
        For k = 1 To n_forn  
            EnerSupAcc = EnerSupAcc + EnerSup(j, OrdSup(j, k))
            If supply > EnerSupAcc Then
                AlocateSup(j, OrdSup(j, k)) = EnerSup(j, OrdSup(j, k))
            Else
                FracEnerSup = supply - (EnerSupAcc - EnerSup(j, OrdSup(j, k)))
                AlocateSup(j, OrdSup(j, k)) = FracEnerSup
                Exit For
            End If
        Next k
    Next j
    
    
    For i = 1 To n_slot
        For j = 1 To n_blocks
            Worksheets("AlocateBlock").Cells(6 + i, 2 + j) = AlocateBlock(i, j)
        Next j
    Next i
    
    For i = 1 To n_slot
        For j = 1 To n_blocks
            Worksheets("AlocateEnerLoad").Cells(6 + i, 2 + j) = AlocateEnerLoad(i, j)
        Next j
    Next i
    
    For i = 1 To n_slot
        For j = 1 To n_forn
            Worksheets("AlocateSup").Cells(6 + i, 2 + j) = AlocateSup(i, j)
        Next j
    Next i
    sumsG(iterTotalN) = Application.Sum(Range(Worksheets("AlocateSup").Cells(7, 3), Worksheets("AlocateSup").Cells(30, 3)))
Next iterTotalN
Worksheets("AlocateSup").Cells(31, 3) = Application.Sum(sumsG)
End Sub
