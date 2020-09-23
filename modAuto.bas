Attribute VB_Name = "modAuto"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit
Const mSolitaireFundationRule = 1

Const mCardAce = 1
Const mKingCard = 13

Const mTablePile = 1
Const mFundationPile = 2
Const mDiscardPile = 3
Const mDealPile = 4

Function FuncCardMove() As Integer
    CPUPlay = True
    SubCardMoveAnimated gPile, gPileCardsPicts, gCardsInPile
End Function

Sub SubFindInDiscardPile(pPile As PILESINFORMATION)
Dim lPileNumberInPile As Integer
Dim lTotalCardsInDiscardPile As Integer

    lTotalCardsInDiscardPile = FuncGetTotalCardsInDiscardPile(1)
    lPileNumberInPile = 1
    
    SubGetCardInfoFromDiscardPile lPileNumberInPile, pPile
End Sub

Sub SubFindInFundationPile(pCard As CardInfo, pPile As PILESINFORMATION)
'Find in all fundation piles work
Dim i As Integer
    pPile.tnPileNumberInPile = 0
    If pCard.tCardValue = mCardAce Then
        For i = 1 To 4
            SubGetCardInfoFromFundationPile i, pPile
            If 0 = pPile.tCardIndex Then
                Exit Sub
            End If
        Next i
    Else
        For i = 1 To 4
            SubGetCardInfoFromFundationPile i, pPile
            
            If pPile.tCardIndex <> 0 Then
            'if the card can go to the fundation and if exist 1 card in fundation
                If CheckCardsPermissions(pCard, pPile.tnPileInfo.tPileInfoCard, mSolitaireFundationRule) Then
                    Exit Sub
                End If
            End If
        Next i
    End If
    pPile.tnPileNumberInPile = 0
End Sub
Function FuncFoundationValidCardDrop(pCard As CardInfo, pCardDest As CardInfo) As Integer
    If pCardDest.tCardSuit = 0 Then
        FuncFoundationValidCardDrop = (pCard.tCardValue = mCardAce)
    Else
    'exist cards in the fundationpile
        FuncFoundationValidCardDrop = CheckCardsPermissions(pCard, pCardDest, mSolitaireFundationRule)
    End If
End Function

Sub SubDealPileFindPlaceForCard()
'When the player click on card in DealPile and the automove is true
'this function search for a valid pile
Dim lReturn As Integer
    
    'First search in fundation pile
    SubFindInFundationPile gPile.tnPileInfo.tPileInfoCard, gPileCardsPicts
    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
        gCardsInPile = 1
        lReturn = FuncCardMove()
        Exit Sub
    End If
    'and in Table pile
    SubFindInTablePile gPile, gPileCardsPicts
    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
        gCardsInPile = 1
        lReturn = FuncCardMove()
        Exit Sub
    End If
    'and last in discard pile
    SubFindInDiscardPile gPileCardsPicts
    If gPileCardsPicts.tnPileNumberInPile <> 0 Then
        gCardsInPile = 1
        lReturn = FuncCardMove()
        Exit Sub
    End If
End Sub

Sub SubNewGame()
    SubRedimArrays
    SubGetRuleInfo
    SubResetStats
    frmTable.Caption = GameSolitaire
End Sub


Sub SubResize()
    SubRedimPilesAndRefreshForm
    SubRefreshPilesBackground True
End Sub

Function SubSetFaceUp(pPile As PILESINFORMATION) As Integer
Dim lTotalCardsInPile As Integer
Const lMoveable = True
    
    
    SubSetFaceUp = False
    Select Case pPile.tnPileNumber
    Case mTablePile
        lTotalCardsInPile = SubHowManyCardsInPile(pPile.tnPileNumberInPile)
        If lTotalCardsInPile = pPile.tCardIndex Then 'is the first card in pile?
            SubTurnFaceUpTable pPile.tnPileNumberInPile, pPile.tCardIndex, lMoveable, gSaveUndo
            SubUpdateMobealble pPile
            SubSetFaceUp = True
        End If
    Case mDealPile
        lTotalCardsInPile = SubGetHowManyCardsLeftInPile(pPile.tnPileNumberInPile)
        If lTotalCardsInPile = pPile.tCardIndex Then 'is the first card in pile?
            SubTurnFaceUpDeal pPile.tnPileNumberInPile, pPile.tCardIndex, lMoveable, gSaveUndo
            SubUpdateMobealble pPile
            SubSetFaceUp = True
        End If
    End Select

End Function


Sub SubUpdateMobealbleToAll()
Dim lPile As PILESINFORMATION
Dim i As Integer
    
    lPile.tnPileNumber = mTablePile
    For i = 1 To 7
        lPile.tnPileNumberInPile = i
        SubUpdateMobealble lPile
    Next i
    
    lPile.tnPileNumber = mFundationPile
    For i = 1 To 4
        lPile.tnPileNumberInPile = i
        SubUpdateMobealble lPile
    Next i
    
    lPile.tnPileNumber = mDiscardPile
        lPile.tnPileNumberInPile = 1
        SubUpdateMobealble lPile
    
    lPile.tnPileNumber = mDealPile
        lPile.tnPileNumberInPile = 1
        SubUpdateMobealble lPile
End Sub

Sub SubUpdateMobealble(pPile As PILESINFORMATION)
    Select Case pPile.tnPileNumber
    
    Case mTablePile
        If SubHowManyCardsInPile(pPile.tnPileNumberInPile) > 0 Then
            SubSetMoveableToTablePile pPile.tnPileNumberInPile
        End If
        
    Case mFundationPile
        SubSetMoveableToFundationPile pPile.tnPileNumberInPile, True
    
    Case mDiscardPile
        SubGetCardInfoFromDiscardPile pPile.tnPileNumberInPile, pPile
        If pPile.tCardIndex > 0 Then
            SubSetMoveableToDiscardPile pPile.tnPileNumberInPile, pPile.tCardIndex, True
        End If
    
    Case mDealPile
        SubGetCardInfoFromDealPile pPile.tnPileNumberInPile, pPile
        If pPile.tCardIndex > 0 Then
            SubSetMoveableToDealPile pPile.tnPileNumberInPile, pPile.tCardIndex, pPile.tnPileInfo.tPileCardBlocked
        End If
    
    Case Else
    End Select
End Sub


Function SubTablePileFindPlaceForCard(pPile As PILESINFORMATION, pPileDest As PILESINFORMATION, ByVal pCardsInPile As Integer) As Integer
'this func check where player as droped the card and check if is a valid solitaire rule
    If pCardsInPile < 1 Then
        SubTablePileFindPlaceForCard = False
        Exit Function
    End If
    
    If pPileDest.tnPileNumber = 0 Or pPile.tnPileNumber = 0 Then
        SubTablePileFindPlaceForCard = False
        Exit Function
    End If
    
    If pPile.tnPileNumberInPile = 0 Or pPileDest.tnPileNumberInPile = 0 Then
        SubTablePileFindPlaceForCard = False
        Exit Function
    End If
    
    Select Case pPileDest.tnPileNumber
    Case mTablePile
        If pPile.tnPileNumber = mTablePile And pPile.tnPileNumberInPile = pPileDest.tnPileNumberInPile Then
            SubTablePileFindPlaceForCard = False
            Exit Function
        End If
        If pPileDest.tnPileInfo.tPileInfoCard.tCardSuit = 0 Then
            SubTablePileFindPlaceForCard = pPile.tnPileInfo.tPileInfoCard.tCardValue = mKingCard
        Else
        'all ok, see only if is a valid solitaire rule
            SubTablePileFindPlaceForCard = CheckCardsPermissions(pPileDest.tnPileInfo.tPileInfoCard, pPile.tnPileInfo.tPileInfoCard, gSolitaireTableRule)
        End If
        Exit Function
        
    Case mFundationPile
        If pPile.tnPileNumber = mFundationPile Then
            SubTablePileFindPlaceForCard = False
            Exit Function
        End If
        If pCardsInPile = 1 Then
            SubTablePileFindPlaceForCard = FuncFoundationValidCardDrop(pPile.tnPileInfo.tPileInfoCard, pPileDest.tnPileInfo.tPileInfoCard)
        Else
        'if total of cards > 1 is a invalid drop for fundation pile
            SubTablePileFindPlaceForCard = False
        End If
    Case mDiscardPile
        SubTablePileFindPlaceForCard = (pPile.tnPileNumber = mDealPile)
    Case mDealPile
        SubTablePileFindPlaceForCard = False
    End Select
End Function

Sub SubFindInTablePile(pPile As PILESINFORMATION, pPileDest As PILESINFORMATION)
Dim i As Integer
    
    'quando existe já uma carta por baixo do pilha table(auto)
    If pPile.tnPileInfo.tPileInfoCard.tCardValue <> mKingCard Then
        For i = 1 To 7
            If Not (mTablePile = pPile.tnPileNumber And i = pPile.tnPileNumberInPile) Then
                SubGetCardInfoFromTablePile i, pPileDest
                If pPileDest.tCardIndex <> 0 Then
                    If pPileDest.tnPileInfo.tPileCardBlocked Then
                        If CheckCardsPermissions(pPileDest.tnPileInfo.tPileInfoCard, pPile.tnPileInfo.tPileInfoCard, gSolitaireTableRule) Then
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next i
    End If
    
    'quando não existe nenhuma carta por baixo da pilha table(auto)
    If pPile.tnPileInfo.tPileInfoCard.tCardValue = mKingCard Then
        For i = 1 To 7
            SubGetCardInfoFromTablePile i, pPileDest
            If 0 = pPileDest.tCardIndex Then
                Exit Sub
            End If
        Next i
    End If
    pPileDest.tnPileNumberInPile = 0
End Sub

