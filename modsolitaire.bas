Attribute VB_Name = "modsolitaire"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 23/09/2002
'**********************************************************************************



Option Explicit
Const mSaveUndo = True
Const mKingCard = 13
Const mSolitaireRule = 2
Const ToolBarHeight = 20
Const MFaceUp = True
Const MFaceDown = False
Const mPileDealNumber = 1 ' number of pile in DealPile



Function SubCheckIfCompleted() As Variant
'Check if all fundation piles is full of cards, if yes the game is completed
Dim i As Integer
Dim lPile As PILESINFORMATION
    For i = 1 To 4
        SubGetCardInfoFromFundationPile i, lPile
        If lPile.tnPileInfo.tPileInfoCard.tCardSuit = 0 Or lPile.tnPileInfo.tPileInfoCard.tCardValue <> mKingCard Then
            SubCheckIfCompleted = False
            Exit Function
        End If
    Next
    SubCheckIfCompleted = True
End Function

Function FuncCpuFindMove() As Integer
'This function is brain of CPU for play Solitaire alone
Const lnPileNumberInPile = 1
Const lMoveable = True

    'first check if exist a valid move from tablepile
    SubFindInTablePile gPile, gPileCardsPicts
    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
        gCardsInPile = 1
        FuncCpuFindMove = Not FuncCardMove()
        Exit Function 'deal card from discardpile to tablepile
    End If
    
    'check for refil dealpile
    SubGetCardInfoFromDealPile lnPileNumberInPile, gPile
    If gPile.tCardIndex = 0 Then 'no more cards in deal pile?
        If FuncGetTotalCardsInDiscardPile(1) <> 0 Then 'and in discard pile?
            SubCardDealRefil gPile.tnPileNumberInPile, 1 'yes, go to refill func
            SubGetCardInfoFromDealPile lnPileNumberInPile, gPile
            FuncCpuFindMove = True
            Exit Function
        Else
            FuncCpuFindMove = False
            Exit Function 'no more cards in dealpile and discardpile
        End If
    End If
    
    
    'and in the table pile again
    SubFindInTablePile gPile, gPileCardsPicts
    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
        gCardsInPile = 1
        FuncCpuFindMove = Not FuncCardMove()
        Exit Function 'deal card from dealpile to tablepile
    End If
    
    
    
    SubGetCardInfoFromDealPile lnPileNumberInPile, gPile
    SubFindInDiscardPile gPileCardsPicts
    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
        gCardsInPile = 1
        FuncCpuFindMove = Not FuncCardMove()
        Exit Function 'deal cart from dealpile to discardpile
    End If
    FuncCpuFindMove = False
    
End Function

Function FuncAutoFaceUpDealPile() As Integer
Dim pPlacePileInPile
Const lPileDeal = 1 'the only pile in DealPile
Const lMoveable = True

    
    For pPlacePileInPile = 1 To 7
        SubGetCardInfoFromTablePile pPlacePileInPile, gPile
        If gPile.tCardIndex <> 0 Then
            If Not gPile.tnPileInfo.tPileCardBlocked Then
                SubTurnFaceUpTable pPlacePileInPile, gPile.tCardIndex, lMoveable, mSaveUndo
                SubSetCardMoveable pPlacePileInPile, gPile.tCardIndex, lMoveable
                FuncAutoFaceUpDealPile = True
                Exit Function
            End If
        End If
    Next pPlacePileInPile
    
    SubGetCardInfoFromDealPile lPileDeal, gPile
    If gPile.tCardIndex <> 0 Then 'any cards in deal pile?
        If Not gPile.tnPileInfo.tPileCardBlocked Then
            SubTurnFaceUpDeal gPile.tnPileNumberInPile, gPile.tCardIndex, lMoveable, mSaveUndo
            SubSetMoveableToDealPile gPile.tnPileNumberInPile, gPile.tCardIndex, lMoveable
            FuncAutoFaceUpDealPile = True
            Exit Function
        End If
    End If
    
    
    FuncAutoFaceUpDealPile = False
End Function

Function FuncAutoMoveLayout() As Integer
Dim lPileNumberInPile As Integer
Dim lNumberOfUndos As Integer
Dim lPileNumberInPileTMP As Integer

    lPileNumberInPile = 1
    lNumberOfUndos = 1
    lPileNumberInPileTMP = lPileNumberInPile
    Do
        SubGetPileInfo lPileNumberInPile, gPile
        If gPile.tCardIndex <> 0 Then
            If Not (gPile.tCardIndex = 1 And gPile.tnPileInfo.tPileInfoCard.tCardValue = mKingCard) Then
                SubFindInTablePile gPile, gPileCardsPicts
                If gPileCardsPicts.tnPileNumberInPile <> 0 Then
                    gCardsInPile = SubHowManyCardsInPile(lPileNumberInPile) - gPile.tCardIndex + 1
                    FuncAutoMoveLayout = Not FuncCardMove()
                    Exit Function
                End If
            End If
        End If
        SubActualizeUndosNumber lPileNumberInPile, lNumberOfUndos, 7
        If lPileNumberInPile = lPileNumberInPileTMP Then Exit Do
    Loop
    FuncAutoMoveLayout = False
End Function

Function FuncAutoMoveToFoundation() As Integer
Dim i As Integer
Const lnPileNumberInPile = 1
    'first search in table pile for valid moves
    For i = 1 To 7
        SubGetCardInfoFromTablePile i, gPile
        If gPile.tCardIndex <> 0 Then
            If gPile.tnPileInfo.tPileCardBlocked Then
                SubFindInFundationPile gPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
                If gPileCardsPicts.tnPileNumberInPile <> 0 Then
                    gCardsInPile = 1
                    FuncAutoMoveToFoundation = Not FuncCardMove()
                    Exit Function
                End If
            End If
        End If
    Next i
    
    'next search in discardpile
        SubGetCardInfoFromDiscardPile 1, gPile
        SubFindInFundationPile gPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
        If gPileCardsPicts.tnPileNumberInPile <> 0 Then
            gCardsInPile = 1
            FuncAutoMoveToFoundation = Not FuncCardMove()
            Exit Function
        End If
        
    'and finaly search in dealpile
    SubGetCardInfoFromDealPile lnPileNumberInPile, gPile
    If gPile.tCardIndex <> 0 Then
        If gPile.tnPileInfo.tPileCardBlocked Then
            SubFindInFundationPile gPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
            If gPileCardsPicts.tnPileNumberInPile <> 0 Then
                gCardsInPile = 1
                FuncAutoMoveToFoundation = Not FuncCardMove()
                Exit Function
            End If
        End If
    End If
    
    
    
    FuncAutoMoveToFoundation = False
End Function

Function FuncFastWin(pFastWin As Integer) As Variant
'check if fast win is true
Dim i As Integer
Dim j As Integer
Dim lCard As CardInfo
Dim lCardDest As CardInfo

    If Not pFastWin Then 'if the fast win is disable, check if game is completed
        FuncFastWin = SubCheckIfCompleted()
        Exit Function
    End If
    'check in all table piles if exist any card blocked
    For i = 1 To 7
        For j = 1 To SubHowManyCardsInPile(i)
            If Not FuncGetMoveableFromTablePile(i, j) Then
                FuncFastWin = False
                Exit Function
            End If
            If j = 1 Then
                SubGetInfoFwin i, j, lCard
            Else
                SubGetInfoFwin i, j, lCardDest
                lCard = lCardDest
            End If
        Next j
    Next i
    FuncFastWin = True
End Function

Private Sub SubGetCordsForDrawCards(pTablePile() As tCordinates, pFundationPile() As tCordinates, pDiscardPilePile() As tCordinates, pDealPilePile() As tCordinates)
Dim i As Integer
Dim lToolBarDiference As Integer
Dim lBord As Integer
Dim lSevenPlaces As Integer
Dim lCardMeasure As tCordinates
Dim lLeftSpace As Variant

    SubGetCardMeasure lCardMeasure
    lSevenPlaces = frmTable.ScaleWidth / 7
    lBord = (lSevenPlaces - lCardMeasure.tX) / 2
    lLeftSpace = lBord + lSevenPlaces * 3
    
    For i = 1 To 4
        pFundationPile(i).tX = lLeftSpace + (i - 1) * lSevenPlaces
        pFundationPile(i).tY = ToolBarHeight
    Next i
    
    lToolBarDiference = ToolBarHeight + lCardMeasure.tY + 10
    
    For i = 1 To 7
        pTablePile(i).tX = lBord + (i - 1) * lSevenPlaces
        pTablePile(i).tY = lToolBarDiference
    Next i
    
    pDealPilePile(1).tX = lBord
    pDealPilePile(1).tY = ToolBarHeight
    pDiscardPilePile(1).tX = lBord + lSevenPlaces
    pDiscardPilePile(1).tY = ToolBarHeight
End Sub

Sub SubRefillDealPile(pPile As PILESINFORMATION)
    SubCardDealRefil pPile.tnPileNumberInPile, 1
End Sub

Sub SubRedimArrays()
Dim lTablePile() As tCordinates
Dim lFundationPile() As tCordinates
Dim lDiscardPilePile() As tCordinates
Dim lDealPilePile() As tCordinates

    ReDim lTablePile(1 To 7) As tCordinates
    ReDim lFundationPile(1 To 4) As tCordinates
    ReDim lDiscardPilePile(1) As tCordinates
    ReDim lDealPilePile(1) As tCordinates
    
    gBmpFile = App.Path + "\Cards.bmp"
    SubLoadBmpFile gBmpFile
    
    SubGetCordsForDrawCards lTablePile(), lFundationPile(), lDiscardPilePile(), lDealPilePile()
    SubInitPiles lTablePile(), lFundationPile(), lDiscardPilePile(), lDealPilePile()
    
End Sub

Sub SubGetRuleInfo()
Dim lCords As tCordinates
    SubGetCardMeasure lCords 'get card height and width pixels
    gSolitaireTableRule = mSolitaireRule
End Sub


Sub SubDealCardsInTableGame()
Dim i As Integer
Dim j As Integer
Dim lCard As CardInfo

    For i = 1 To 7 ' 7 piles
        For j = i To 7 'max cards in pile 7
            SubAddCardToDeck lCard
            If j = i Then
                SubPutCardsInTablePile lCard, j, MFaceUp, True
            Else
                SubPutCardsInTablePile lCard, j, MFaceDown, True
            End If
        Next j
    Next i
    
    ' the rest of cards go to deal pile
    Do
        SubAddCardToDeck lCard
        If lCard.tCardSuit = 0 Then
            Exit Do
        End If
        SubPutCardsInDealPile lCard, mPileDealNumber, MFaceDown, True
    Loop
End Sub

Sub SubRedimPilesAndRefreshForm()
Dim lTablePile() As tCordinates
Dim lFundationPile() As tCordinates
Dim lDiscardPilePile() As tCordinates
Dim lDealPilePile() As tCordinates

    ReDim lTablePile(1 To 7) As tCordinates
    ReDim lFundationPile(1 To 4) As tCordinates
    ReDim lDiscardPilePile(1) As tCordinates
    ReDim lDealPilePile(1) As tCordinates
    
    SubGetCordsForDrawCards lTablePile(), lFundationPile(), lDiscardPilePile(), lDealPilePile()
    SubSetCordsForPiles lTablePile(), lFundationPile(), lDiscardPilePile(), lDealPilePile()
End Sub
