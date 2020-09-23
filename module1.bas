Attribute VB_Name = "Module1"
'**********************************************************************************
'**Solitaire 1.0
'**Copyright by Paulo Matoso
'**E-Mail: paulomt1@clix.pt
'**
'**
'**Last Modification ---> 17/08/2002
'**********************************************************************************



Option Explicit
Dim mSuit(1 To 4) As Integer ' 4 naipes
Dim mForm As Form
Dim mBackCardPicture As Integer
Dim mCardWith As Integer
Dim mCardHeight As Integer
Dim mCard As CardInfo

Dim mTablePileInfo() As tPileInfo
Dim mFundationPileInfo() As tPileInfo
Dim mDiscardPileInfo() As tPileInfo
Dim mDealPileInfo() As tPileInfo

Dim mCardSelectedInPile() As Integer
Dim mDiscardPileCards() As Integer
Dim mHowManyCardsLeft() As Integer ' return how many cards have i (pile)

Dim mDeck(1 To 52) As CardInfo
Dim mTotalCards As Integer
Const mMaxUndosPermited = 50
Dim mUndoBuffer(1 To mMaxUndosPermited) As UNDOTYPE
Dim mUndosLeft As Integer
Dim mTempUndos As Integer


Dim mAnimationSpeed As Single
Dim mAnimationInterval As Single
Dim noCard As Integer ' if the player clicked on green this value is false
                    ' else if player clicked on a card this value is true
Dim mPile As PILESINFORMATION
Dim mPileDest As PILESINFORMATION

Dim mLocalizationOnCardClicked As tCordinates
Dim nCardsDroped As Integer
Dim mCardPixels As tCordinates
Dim mScrunchValue As Integer
Dim mCalculedSpeed As Single
Dim mCardStartCordinatesClicked As tCordinates
Dim mSpeedFactor As Single

Const SRCCOPY = &HCC0020

Sub SubPutCardsInDealPile(pCard As CardInfo, ByVal pnPileNumberInPile As Integer, ByVal pMoveable As Integer, ByVal pRefill As Integer)
Dim lnTotalCardsInPile As Integer
Dim lReturn As Integer
'draw cards into dealpile

    If gSoundTurnedOn Then
        lReturn = sndPlaySound(gSoundDeal, SND_ASYNC)
    End If
    
    lnTotalCardsInPile = mHowManyCardsLeft(pnPileNumberInPile) + 1 ' add cards to deal pile
    mHowManyCardsLeft(pnPileNumberInPile) = lnTotalCardsInPile
    mDealPileInfo(pnPileNumberInPile, lnTotalCardsInPile).tPileInfoCard = pCard
    mDealPileInfo(pnPileNumberInPile, lnTotalCardsInPile).tPileCardBlocked = pMoveable
    
    If pRefill Then
        If lnTotalCardsInPile = 1 Then
        ' if the first card to add to deal pile
            SubDrawCardInPile mDealPileInfo(pnPileNumberInPile, lnTotalCardsInPile), True, mCard, 0
        Else
            SubDrawCardInPile mDealPileInfo(pnPileNumberInPile, lnTotalCardsInPile), True, mDealPileInfo(pnPileNumberInPile, lnTotalCardsInPile - 1).tPileInfoCard, 0
        End If
    End If
End Sub

Sub SubPutCardsInDiscardPile(pCard As CardInfo, ByVal pPileNumberInPile As Integer, ByVal pMoveable As Integer, ByVal pUpdate As Integer)
Dim lTotalCardsInDiscardPile As Integer
Dim lReturn As Integer
    If gSoundTurnedOn Then
        lReturn = sndPlaySound(gSoundDeal, SND_ASYNC)
    End If
    lTotalCardsInDiscardPile = mDiscardPileCards(pPileNumberInPile) + 1
    mDiscardPileCards(pPileNumberInPile) = lTotalCardsInDiscardPile
    mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile).tPileInfoCard = pCard
    mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile).tPileCardBlocked = pMoveable
    
    If pUpdate Then 'when deal from tablepile to discard pile(only when player undo last move)
    'this update backgownd card in table pile
        If lTotalCardsInDiscardPile = 1 Then
            SubDrawCardInPile mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile), True, mCard, 0
        Else
            SubDrawCardInPile mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile), True, mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile - 1).tPileInfoCard, 0
        End If
    End If
End Sub

Private Sub SubPutCardsInFundationPile(pCard As CardInfo, pPileInPileNumber As Integer, ByVal pUpdate As Integer)
    mFundationPileInfo(pPileInPileNumber).tPileInfoCard = pCard
    mFundationPileInfo(pPileInPileNumber).tPileCardBlocked = True
    
    If pUpdate Then
        SubDrawCardInPile mFundationPileInfo(pPileInPileNumber), True, mCard, 0
    End If
End Sub

Sub SubPutCardsInTablePile(pCard As CardInfo, ByVal pPileInPile As Integer, ByVal pMoveable As Integer, ByVal pInitGame As Integer)
Dim lCardsInPile As Integer
Dim lReturn As Integer
    lCardsInPile = mCardSelectedInPile(pPileInPile) + 1
    mCardSelectedInPile(pPileInPile) = lCardsInPile
    mTablePileInfo(pPileInPile, lCardsInPile).tPileInfoCard = pCard
    mTablePileInfo(pPileInPile, lCardsInPile).tPileCardBlocked = pMoveable
    
    If pInitGame Then
        If lCardsInPile = 1 Then
            SubDrawCardInPile mTablePileInfo(pPileInPile, lCardsInPile), True, mCard, 18
        Delay 200
        If gSoundTurnedOn Then
            lReturn = sndPlaySound(gSoundDrawUp, SND_ASYNC)
        End If
        Else
            SubDrawCardInPile mTablePileInfo(pPileInPile, lCardsInPile), True, mTablePileInfo(pPileInPile, lCardsInPile - 1).tPileInfoCard, 18
        Delay 200
        If gSoundTurnedOn Then
            lReturn = sndPlaySound(gSoundDrawUp, SND_ASYNC)
        End If
        End If
    End If
End Sub

Private Function FuncSaveMove(pPileSource As Integer, ByVal pPileInPileSource As Integer, pPileDest As Integer, ByVal pPileInPileDest As Integer, ByVal pTotalCardsMoved As Integer) As Integer

Dim lTotalCardsInPile As Integer
Dim lValue As Integer
Dim lUndosLeft As Integer

    Select Case pPileDest
    Case TablePile
        lTotalCardsInPile = mCardSelectedInPile(pPileInPileDest)
    Case FundationPile
        lTotalCardsInPile = mFundationPileInfo(pPileInPileDest).tPileInfoCard.tCardValue
    Case DiscardPile
        lTotalCardsInPile = mDiscardPileCards(pPileInPileDest)
    Case DealPile
        lTotalCardsInPile = mHowManyCardsLeft(pPileInPileDest)
    End Select
    
    mUndoBuffer(mUndosLeft).tUndoPileSource = pPileSource 'pile number source
    mUndoBuffer(mUndosLeft).tUndoPileInPileSource = pPileInPileSource 'pile in pile source
    mUndoBuffer(mUndosLeft).tUndoPileDest = pPileDest 'pile number dest
    mUndoBuffer(mUndosLeft).tUndoPileInPileDest = pPileInPileDest 'pile in pile number dest
    mUndoBuffer(mUndosLeft).tUndoCardsInPileAfterMove = lTotalCardsInPile 'number of cards in pile dest
    mUndoBuffer(mUndosLeft).tUndoTotalCardsMoved = pTotalCardsMoved ' how many cards move
    
    lUndosLeft = mUndosLeft
    SubActualizeUndosNumber lUndosLeft, -1, mMaxUndosPermited
    lValue = FuncUpdateUndoPiles(mUndoBuffer(mUndosLeft), mUndoBuffer(lUndosLeft))
    SubActualizeUndosNumber lUndosLeft, -1, mMaxUndosPermited
    FuncSaveMove = lValue Or FuncUpdateUndoPiles(mUndoBuffer(mUndosLeft), mUndoBuffer(lUndosLeft))
    SubActualizeUndosNumber mUndosLeft, 1, mMaxUndosPermited
    
    If mUndosLeft = mTempUndos Then
        SubActualizeUndosNumber mTempUndos, 1, mMaxUndosPermited
    End If
    
End Function

Sub SubPlayApplause()
Dim lRet As Integer
    If gSoundTurnedOn Then
        lRet = sndPlaySound(gSoundApplause, SND_ASYNC)
    End If
End Sub




Sub SubGetCardMeasure(pCardMeasure As tCordinates)
'Return the card Width and Height in pixels
    pCardMeasure.tX = mCardWith
    pCardMeasure.tY = mCardHeight
End Sub

Sub SubSetCordsForPiles(pTablePile() As tCordinates, pFundationPile() As tCordinates, pDiscardPilePile() As tCordinates, pDealPilePile() As tCordinates)
Dim i As Integer
Dim j As Integer
Dim lNoScrunch As Integer 'No scrunch for TablePile and DealPile

    For i = 1 To 7
        mTablePileInfo(i, 1).tPileInfoCords = pTablePile(i)
        For j = 1 To UBound(mTablePileInfo, 2)
            mTablePileInfo(i, j).tPileInfoCords.tX = mTablePileInfo(i, 1).tPileInfoCords.tX
            mTablePileInfo(i, j).tPileInfoCords.tY = mTablePileInfo(i, 1).tPileInfoCords.tY + (j - 1) * 18
        Next j
    Next i
    
    For j = 1 To 4
        mFundationPileInfo(j).tPileInfoCords = pFundationPile(j)
    Next j
    
        mDiscardPileInfo(1, 1).tPileInfoCords = pDiscardPilePile(1)
        For j = 1 To UBound(mDiscardPileInfo, 2)
            mDiscardPileInfo(1, j).tPileInfoCords.tX = mDiscardPileInfo(1, 1).tPileInfoCords.tX
            mDiscardPileInfo(1, j).tPileInfoCords.tY = mDiscardPileInfo(1, 1).tPileInfoCords.tY + (j - 1) * lNoScrunch
        Next j
    
        mDealPileInfo(1, 1).tPileInfoCords = pDealPilePile(1)
        For j = 1 To UBound(mDealPileInfo, 2)
            mDealPileInfo(1, j).tPileInfoCords.tX = mDealPileInfo(1, 1).tPileInfoCords.tX
            mDealPileInfo(1, j).tPileInfoCords.tY = mDealPileInfo(1, 1).tPileInfoCords.tY + (j - 1) * lNoScrunch
        Next j
End Sub




Private Function FuncUpdateUndoPiles(pUndoSource As UNDOTYPE, pUndoDest As UNDOTYPE) As Integer
Dim lValue As Integer
    lValue = pUndoSource.tUndoPileSource = pUndoDest.tUndoPileSource
    lValue = lValue And pUndoSource.tUndoPileInPileSource = pUndoDest.tUndoPileInPileSource
    lValue = lValue And pUndoSource.tUndoPileDest = pUndoDest.tUndoPileDest
    lValue = lValue And pUndoSource.tUndoPileInPileDest = pUndoDest.tUndoPileInPileDest
    lValue = lValue And pUndoSource.tUndoCardsInPileAfterMove = pUndoDest.tUndoCardsInPileAfterMove
    FuncUpdateUndoPiles = lValue
End Function



Function SubGetHowManyCardsLeftInPile(ByVal pPileInPile As Integer) As Integer
'return total of cards in pPileInPile var
    SubGetHowManyCardsLeftInPile = mHowManyCardsLeft(pPileInPile)
End Function


Sub SubCardDealRefil(ByVal pnPileNumberInPile As Integer, ByVal pPileNumberInPile As Integer)
Dim i As Integer
Dim lReturn As Integer
Dim lCard As CardInfo

    If mDiscardPileCards(pPileNumberInPile) <> 0 Then 'if no more cards left in discard pile
                                            'don´t play sound
        If gSoundTurnedOn Then
            lReturn = sndPlaySound(gSoundStack, SND_ASYNC) 'refill
        End If
        Delay 300 'for sound
        gNoMoreMoves = gNoMoreMoves + 1
        If gNoMoreMoves = 3 And gMoveMultipleAuto Then
            MsgBox "Sorry, i can´t find more valid moves", MB_ICONASTERISK
            frmTable.ResetAutoPlayVars
        End If
    End If
    
    'Move all the cards in discard pile to Deal Pile
    If mDiscardPileCards(pPileNumberInPile) <> 0 Then
        For i = FuncGetTotalCardsInDiscardPile(pPileNumberInPile) To 1 Step True
            SubGetCardFromDiscardPile pPileNumberInPile, True, lCard
            SubPutCardsInDealPile lCard, pnPileNumberInPile, gFaceDown, True
        Next i
    End If
    mUndosLeft = mTempUndos
End Sub

Sub SubTurnFaceUpDeal(ByVal pPileInPileSource As Integer, ByVal pTotalCardsMoved As Integer, ByVal pMoveable As Integer, ByVal pSaveMove As Integer)
'the deal pile
Dim lReturn As Integer

    If gSoundTurnedOn Then
        lReturn = sndPlaySound(gSoundTurn, SND_ASYNC) 'virar dealpile
    End If
    
    mDealPileInfo(pPileInPileSource, pTotalCardsMoved).tPileCardBlocked = pMoveable
    SubDrawCardInPile mDealPileInfo(pPileInPileSource, pTotalCardsMoved), False, mCard, 0
    
    If pSaveMove Then
        lReturn = FuncSaveMove(DealPile, pPileInPileSource, DealPile, pPileInPileSource, pTotalCardsMoved)
    End If
End Sub



Function FuncGetTotalCardsInDiscardPile(ByVal pPileNumberInPile As Integer) As Integer
    FuncGetTotalCardsInDiscardPile = mDiscardPileCards(pPileNumberInPile)
End Function




Function CheckForClickInCard() As Integer
'this function say if player clicked on a card or on another place
' if player clicked on a card this function return a value true
    CheckForClickInCard = noCard
End Function

Function CheckIfCardAnimationIsFinish() As Integer
Dim lScreenCords As tCordinates
Dim lSpeed As Single
    mSpeedFactor = mSpeedFactor + mAnimationInterval
    lSpeed = mSpeedFactor / mCalculedSpeed
    
    If lSpeed > 1 Then
    'card move is end
        lSpeed = 1
        CheckIfCardAnimationIsFinish = True
        SubActualizePositionOfCardInDest
        SubSaveVarsForUndoRefreshCard
    Else
    'card is moving
        CheckIfCardAnimationIsFinish = False
        lScreenCords.tX = mPile.tnPileInfo.tPileInfoCords.tX + lSpeed * (mPileDest.tnPileInfo.tPileInfoCords.tX - mPile.tnPileInfo.tPileInfoCords.tX)
        lScreenCords.tY = mPile.tnPileInfo.tPileInfoCords.tY + lSpeed * (mPileDest.tnPileInfo.tPileInfoCords.tY - mPile.tnPileInfo.tPileInfoCords.tY)
        SubRefreshCard lScreenCords
    End If
End Function

Sub SubCardMoveAnimated(pPile As PILESINFORMATION, pPileDest As PILESINFORMATION, ByVal pDropedCards As Integer)
    mPile = pPile
    mPileDest = pPileDest
    SubGetStartCrunch 'get scrunch diference
    nCardsDroped = pDropedCards
    mScrunchValue = FuncGetScrunchValue(pPile.tnPileNumber) 'get the scrunch value in destination pile
    mSpeedFactor = 0
    mCalculedSpeed = FuncCalcSpeedAnimation(mPile.tnPileInfo.tPileInfoCords, mPileDest.tnPileInfo.tPileInfoCords) / mAnimationSpeed
    mLocalizationOnCardClicked.tX = 0
    mLocalizationOnCardClicked.tY = 0
    SubDrawCard
End Sub
Private Function FuncCalcSpeedAnimation(pSource As tCordinates, pDest As tCordinates) As Single
    FuncCalcSpeedAnimation = Sqr((pSource.tX - pDest.tX) ^ 2 + (pSource.tY - pDest.tY) ^ 2)
End Function

Sub SubRefreshCard(pMouseCords As tCordinates)
'This function refresh the background when player drag a card or cards or when is computer
'play a move(animation move)
Dim lReturn As Integer
Dim lCardStartCordinates As tCordinates
Dim lCardPixelsMoved As tCordinates
Dim lCardPixels As tCordinates
Dim lBitBltSrc As tCordinates
Dim lBitBltCords As tCordinates
Dim lBitBltCords2 As tCordinates

    lCardStartCordinates.tX = pMouseCords.tX - mLocalizationOnCardClicked.tX 'the start of cordinate X of card clicked
    lCardStartCordinates.tY = pMouseCords.tY - mLocalizationOnCardClicked.tY 'the start of cordinate Y of card clicked
    lCardPixelsMoved.tX = lCardStartCordinates.tX - mCardStartCordinatesClicked.tX
    lCardPixelsMoved.tY = lCardStartCordinates.tY - mCardStartCordinatesClicked.tY
    lCardPixels.tX = mCardPixels.tX + Abs(lCardPixelsMoved.tX)
    lCardPixels.tY = mCardPixels.tY + Abs(lCardPixelsMoved.tY)
    frmImgCrdBuffers.ImgCrdDragBuild.Move 2 * mCardWith, 0, lCardPixels.tX, lCardPixels.tY
    
    If lCardPixelsMoved.tX >= 0 Then
        If lCardPixelsMoved.tY >= 0 Then
            lBitBltSrc.tX = mCardStartCordinatesClicked.tX 'card goes down and right
            lBitBltSrc.tY = mCardStartCordinatesClicked.tY
            lBitBltCords.tX = 0
            lBitBltCords.tY = 0
        Else
            lBitBltSrc.tX = mCardStartCordinatesClicked.tX 'card goes up and right or only right
            lBitBltSrc.tY = mCardStartCordinatesClicked.tY + mCardPixels.tY - lCardPixels.tY
            lBitBltCords.tX = 0
            lBitBltCords.tY = lCardPixels.tY - mCardPixels.tY
        End If
    Else
        If lCardPixelsMoved.tY >= 0 Then
            lBitBltSrc.tX = mCardStartCordinatesClicked.tX + mCardPixels.tX - lCardPixels.tX 'card goes down and left or only left
            lBitBltSrc.tY = mCardStartCordinatesClicked.tY
            lBitBltCords.tX = lCardPixels.tX - mCardPixels.tX
            lBitBltCords.tY = 0
        Else
            lBitBltSrc.tX = mCardStartCordinatesClicked.tX + mCardPixels.tX - lCardPixels.tX 'card goes up and left
            lBitBltSrc.tY = mCardStartCordinatesClicked.tY + mCardPixels.tY - lCardPixels.tY
            lBitBltCords.tX = lCardPixels.tX - mCardPixels.tX
            lBitBltCords.tY = lCardPixels.tY - mCardPixels.tY
        End If
    End If
    lBitBltCords2.tX = lBitBltCords.tX + lCardPixelsMoved.tX
    lBitBltCords2.tY = lBitBltCords.tY + lCardPixelsMoved.tY
    lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBuild.hDC, 0, 0, lCardPixels.tX, lCardPixels.tY, mForm.hDC, lBitBltSrc.tX, lBitBltSrc.tY, SRCCOPY)
    lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBuild.hDC, lBitBltCords.tX, lBitBltCords.tY, mCardPixels.tX, mCardPixels.tY, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
    lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, mCardPixels.tX, mCardPixels.tY, frmImgCrdBuffers.ImgCrdDragBuild.hDC, lBitBltCords2.tX, lBitBltCords2.tY, SRCCOPY)
    lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBuild.hDC, lBitBltCords2.tX, lBitBltCords2.tY, mCardPixels.tX, mCardPixels.tY, frmImgCrdBuffers.ImgCrdDrag.hDC, 0, 0, SRCCOPY)
    lReturn = bitblt(mForm.hDC, lBitBltSrc.tX, lBitBltSrc.tY, lCardPixels.tX, lCardPixels.tY, frmImgCrdBuffers.ImgCrdDragBuild.hDC, 0, 0, SRCCOPY)
    mCardStartCordinatesClicked = lCardStartCordinates
End Sub

Private Sub SubActualizePositionOfCardInDest()
Dim lReturn As Integer
Delay 100
    lReturn = bitblt(mForm.hDC, mCardStartCordinatesClicked.tX, mCardStartCordinatesClicked.tY, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleWidth, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleHeight, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
    lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleWidth, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleHeight, mForm.hDC, mPileDest.tnPileInfo.tPileInfoCords.tX, mPileDest.tnPileInfo.tPileInfoCords.tY, SRCCOPY)
    lReturn = bitblt(mForm.hDC, mPileDest.tnPileInfo.tPileInfoCords.tX, mPileDest.tnPileInfo.tPileInfoCords.tY, frmImgCrdBuffers.ImgCrdDrag.ScaleWidth, frmImgCrdBuffers.ImgCrdDrag.ScaleHeight, frmImgCrdBuffers.ImgCrdDrag.hDC, 0, 0, SRCCOPY)
    mCardStartCordinatesClicked = mPileDest.tnPileInfo.tPileInfoCords
End Sub



Private Sub SubSaveVarsForUndoRefreshCard()
'em modo animado
'This function save the va
Dim lReturn As Integer
Dim lCardsInPile As Integer
Dim lCard As CardInfo
Dim lCardDest As CardInfo
Dim i As Integer
Dim lScrunch As Integer
    If gSoundTurnedOn Then
        lReturn = sndPlaySound(gSoundDeal, SND_ASYNC) 'place card in tablepile
    End If
    Delay 300 'wait some time for play 2 sounds
    lCard = mPile.tnPileInfo.tPileInfoCard
    Select Case mPileDest.tnPileNumber
    Case TablePile
        lScrunch = 18
        
        'check if the height of card is great of window
        If mCardStartCordinatesClicked.tY + mCardHeight - lScrunch > mForm.ScaleHeight Then
            lCardDest = mTablePileInfo(mPileDest.tnPileNumberInPile, mCardSelectedInPile(mPileDest.tnPileNumberInPile)).tPileInfoCard
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mCardHeight - lScrunch, frmImgCrdBuffers.pictCrdImage.hDC, (lCardDest.tCardValue - 1) * mCardWith, lCardDest.tCardSuit * mCardHeight + lScrunch, SRCCOPY)
        ElseIf mCardStartCordinatesClicked.tY + mCardHeight > mForm.ScaleHeight Then
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mForm.ScaleHeight - mCardStartCordinatesClicked.tY, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
        Else
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mCardHeight, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
        End If
        'quando são mais que uma carta a mover
        For i = 2 To nCardsDroped
            lCardsInPile = mPile.tCardIndex + i - 1
            lCard = mTablePileInfo(mPile.tnPileNumberInPile, lCardsInPile).tPileInfoCard
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, (lCard.tCardSuit + 1) * mCardHeight - lScrunch, mCardWith, lScrunch, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, mCardHeight + (i - 2) * lScrunch, SRCCOPY)
        Next i
    Case FundationPile
    'quando encaixa na fundation pile(animate mode)
        lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mCardHeight, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
    Case DealPile
        lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mCardHeight, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
    Case DiscardPile
        lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mCardHeight, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
    End Select
End Sub

Private Sub SubDrawCard()
Dim lReturn As Integer
Dim i As Integer
Dim lCardsLeftInPile As Integer
Dim lCard As CardInfo
'só quando é o jogador a arrastar ou a clickar

    mCardStartCordinatesClicked = mPile.tnPileInfo.tPileInfoCords
    mCardPixels.tX = mCardWith
    mCardPixels.tY = mCardHeight + (nCardsDroped - 1) * mScrunchValue
    frmImgCrdBuffers.ImgCrdDrag.Move 0, 0, mCardPixels.tX, mCardPixels.tY
    frmImgCrdBuffers.ImgCrdDragBackGround.Move mCardWith, 0, mCardPixels.tX, mCardPixels.tY
    
    'se for maior do que a janela
    If mCardStartCordinatesClicked.tY + frmImgCrdBuffers.ImgCrdDrag.Height > mForm.ScaleHeight Then
        Select Case mPile.tnPileNumber
        Case TablePile
            For i = 0 To nCardsDroped - 1
                lCard = mTablePileInfo(mPile.tnPileNumberInPile, mPile.tCardIndex + i).tPileInfoCard
                lReturn = bitblt(frmImgCrdBuffers.ImgCrdDrag.hDC, 0, i * mScrunchValue, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
            Next i
        Case FundationPile
            lCard = mFundationPileInfo(mPile.tnPileNumberInPile).tPileInfoCard
            lReturn = bitblt(frmImgCrdBuffers.ImgCrdDrag.hDC, 0, 0, mCardWith, mCardHeight, frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
        Case DiscardPile
            For i = 0 To nCardsDroped - 1
                lCard = mDiscardPileInfo(mPile.tnPileNumberInPile, mPile.tCardIndex + i).tPileInfoCard
                lReturn = bitblt(frmImgCrdBuffers.ImgCrdDrag.hDC, 0, i * mScrunchValue, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
            Next i
        Case DealPile
            For i = 0 To nCardsDroped - 1
                lCard = mDealPileInfo(mPile.tnPileNumberInPile, mPile.tCardIndex + i).tPileInfoCard
                lReturn = bitblt(frmImgCrdBuffers.ImgCrdDrag.hDC, 0, i * mScrunchValue, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
            Next i
        End Select
    Else
    'guarda a carta que se vai mover no buffer(player mode)
        lReturn = bitblt(frmImgCrdBuffers.ImgCrdDrag.hDC, 0, 0, mCardWith, frmImgCrdBuffers.ImgCrdDrag.ScaleHeight, mForm.hDC, mCardStartCordinatesClicked.tX, mCardStartCordinatesClicked.tY, SRCCOPY)
    End If
    
    Select Case mPile.tnPileNumber
    Case TablePile
        For i = nCardsDroped To 1 Step True
            lCardsLeftInPile = mPile.tCardIndex + i - 1 'update total cards in tablepile
            lCard = mTablePileInfo(mPile.tnPileNumberInPile, lCardsLeftInPile).tPileInfoCard
            lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, _
                (i - 1) * mScrunchValue, mCardWith, mCardHeight, _
                frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * _
                mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
        Next i
    Case FundationPile
        lCard = mFundationPileInfo(mPile.tnPileNumberInPile).tPileInfoCard
        lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, _
            mCardWith, mCardHeight, frmImgCrdBuffers.pictBackground.hDC, _
            (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * _
            mCardHeight, SRCCOPY)
    Case DiscardPile
        For i = nCardsDroped To 1 Step True
            lCardsLeftInPile = mPile.tCardIndex + i - 1 'update total cards in discardpile
            lCard = mDiscardPileInfo(mPile.tnPileNumberInPile, lCardsLeftInPile).tPileInfoCard
            'Mostra a carta que está por baixo quando se pega na que está por cima
            lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, _
                (i - 1) * mScrunchValue, mCardWith, mCardHeight, _
                frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) _
                * mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
        Next i
    Case DealPile
        For i = nCardsDroped To 1 Step True
            lCardsLeftInPile = mPile.tCardIndex + i - 1 'update total cards in dealpile
            lCard = mDealPileInfo(mPile.tnPileNumberInPile, lCardsLeftInPile).tPileInfoCard
            'Mostra a carta que está por baixo quando se pega na que está por cima
            lReturn = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, _
               (i - 1) * mScrunchValue, mCardWith, mCardHeight, _
                frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) _
                * mCardWith, lCard.tCardSuit * mCardHeight, SRCCOPY)
        Next i
    End Select
End Sub

Sub SubGetTolerance(pMouseCords As tCordinates, pPile As PILESINFORMATION, ByRef pTolerance As Single)
Dim lCords As tCordinates
    SubRefreshCard pMouseCords
    noCard = False
    lCords.tX = pMouseCords.tX - mLocalizationOnCardClicked.tX + mCardWith / 2
    lCords.tY = pMouseCords.tY - mLocalizationOnCardClicked.tY + mCardHeight / 2
    SubWherePlayerHaveClicked lCords, pPile
    pTolerance = Sqr((pMouseCords.tX - mPile.tnPileInfo.tPileInfoCords.tX - mLocalizationOnCardClicked.tX) ^ 2 + (pMouseCords.tY - mPile.tnPileInfo.tPileInfoCords.tY - mLocalizationOnCardClicked.tY) ^ 2)
End Sub

Sub SubHumanDrawCard(pPile As PILESINFORMATION)
'this function is called only when is the human player
    mPileDest = pPile
    SubGetStartCrunch
    SubActualizePositionOfCardInDest
    SubSaveVarsForUndoRefreshCard
End Sub

Sub SubCheckIfCardIsMoveable(pCords As tCordinates)
Dim lMoveable As Integer
    SubWherePlayerHaveClicked pCords, mPile
    If mPile.tnPileNumber = 0 Or mPile.tCardIndex = 0 Or mPile.tnPileInfo.tPileInfoCard.tCardSuit = mCard.tCardSuit Then
  ' the player have clicked out of a card, or no card in pile
        noCard = False
        Exit Sub
    End If
    Select Case mPile.tnPileNumber
    Case FundationPile
        lMoveable = mFundationPileInfo(mPile.tnPileNumberInPile).tPileInfoMoveable
        nCardsDroped = 1
        mScrunchValue = 0
    Case TablePile
        lMoveable = mTablePileInfo(mPile.tnPileNumberInPile, mPile.tCardIndex).tPileInfoMoveable
        nCardsDroped = mCardSelectedInPile(mPile.tnPileNumberInPile) - mPile.tCardIndex + 1
        mScrunchValue = 18
    Case DealPile
        lMoveable = mDealPileInfo(mPile.tnPileNumberInPile, mPile.tCardIndex).tPileInfoMoveable
        nCardsDroped = 1
        mScrunchValue = 0
    Case DiscardPile
        lMoveable = mDiscardPileInfo(mPile.tnPileNumberInPile, mPile.tCardIndex).tPileInfoMoveable
        nCardsDroped = 1
        mScrunchValue = 0
    End Select
    If lMoveable Then
        noCard = True
        mLocalizationOnCardClicked.tX = pCords.tX - mPile.tnPileInfo.tPileInfoCords.tX
        mLocalizationOnCardClicked.tY = pCords.tY - mPile.tnPileInfo.tPileInfoCords.tY
        SubDrawCard
    Else
        noCard = False
    End If
End Sub

Sub subRefreshPile()
'This function refresh the pile that user have clicked
Dim lRet As Integer
Dim lCords As tCordinates
    lCords = mPile.tnPileInfo.tPileInfoCords
    If nCardsDroped > 0 Then
        lRet = bitblt(mForm.hDC, mCardStartCordinatesClicked.tX, mCardStartCordinatesClicked.tY, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleWidth, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleHeight, frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, SRCCOPY)
        lRet = bitblt(frmImgCrdBuffers.ImgCrdDragBackGround.hDC, 0, 0, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleWidth, frmImgCrdBuffers.ImgCrdDragBackGround.ScaleHeight, mForm.hDC, lCords.tX, lCords.tY, SRCCOPY)
        lRet = bitblt(mForm.hDC, lCords.tX, lCords.tY, frmImgCrdBuffers.ImgCrdDrag.ScaleWidth, frmImgCrdBuffers.ImgCrdDrag.ScaleHeight, frmImgCrdBuffers.ImgCrdDrag.hDC, 0, 0, SRCCOPY)
    End If
    mCardStartCordinatesClicked = lCords
    noCard = False
End Sub

Private Sub SubGetStartCrunch()
        If mPileDest.tCardIndex <> 0 And mPileDest.tnPileNumber = TablePile Then
        'set the start y for draw card(scrunch diference)
            mPileDest.tnPileInfo.tPileInfoCords.tY = mPileDest.tnPileInfo.tPileInfoCords.tY + 18
        End If
End Sub

Function CheckPlayerCardsDroped() As Integer
'this function say how many cards player is droped
    CheckPlayerCardsDroped = nCardsDroped
End Function

Sub SubGetPile(pPile As PILESINFORMATION)
    pPile = mPile
End Sub

Private Sub SubDrawCardsTableBlocked(pCords As tCordinates)
Dim lRet As Integer
'desenha as cartas que estão viradas ao contrario no tablepile
    lRet = bitblt(mForm.hDC, pCords.tX, pCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, (mBackCardPicture - 1) * mCardWith, 0, SRCCOPY)
End Sub

Private Sub SubDrawCardInPile(pPile As tPileInfo, ByVal pBackground As Integer, pCard As CardInfo, ByVal pScrunch As Integer)
'this function draw the background beind cards
Dim lReturn As Integer
Dim lCard As CardInfo
Dim lCords As tCordinates
    lCard = pPile.tPileInfoCard
    lCords = pPile.tPileInfoCords
    
    
    
    If pBackground Then 'draw card bellow the top card
        If lCords.tY + mCardHeight - pScrunch > mForm.ScaleHeight Then
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mCardHeight - pScrunch, frmImgCrdBuffers.pictCrdImage.hDC, (pCard.tCardValue - 1) * mCardWith, pCard.tCardSuit * mCardHeight + pScrunch, SRCCOPY)
        ElseIf lCords.tY + mCardHeight > mForm.ScaleHeight Then
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (lCard.tCardValue - 1) * mCardWith, lCard.tCardSuit * mCardHeight, mCardWith, mForm.ScaleHeight - lCords.tY, mForm.hDC, lCords.tX, lCords.tY, SRCCOPY)
        Else
            lReturn = bitblt(frmImgCrdBuffers.pictBackground.hDC, (pPile.tPileInfoCard.tCardValue - 1) * mCardWith, pPile.tPileInfoCard.tCardSuit * mCardHeight, mCardWith, mCardHeight, mForm.hDC, lCords.tX, lCords.tY, SRCCOPY)
        End If
    End If
    
    If pPile.tPileCardBlocked Then
    'draw cards with face up(refresh mode)

        lReturn = bitblt(mForm.hDC, lCords.tX, lCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, (pPile.tPileInfoCard.tCardValue - 1) * mCardWith, pPile.tPileInfoCard.tCardSuit * mCardHeight, SRCCOPY)
    Else
    'draw cards with face down(refresh mode)
        SubDrawCardsTableBlocked lCords
    End If
    
End Sub

Private Sub SubDrawFundationPlaces(pCords As tCordinates)
Dim lReturn As Long
    lReturn = bitblt(mForm.hDC, pCords.tX, pCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, 10 * mCardWith, 0, SRCCOPY)
End Sub

Private Sub SubDrawTablePlaces(pCords As tCordinates)
Dim lReturn As Long
    lReturn = bitblt(mForm.hDC, pCords.tX, pCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, 11 * mCardWith, 0, SRCCOPY)
End Sub

Private Sub SubDrawDealPlaces(pCords As tCordinates)
Dim lReturn As Long
    lReturn = bitblt(mForm.hDC, pCords.tX, pCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, 9 * mCardWith, 0, SRCCOPY)
End Sub

Private Sub SubDrawDiscardPlaces(pCords As tCordinates)
Dim lReturn As Long
    lReturn = bitblt(mForm.hDC, pCords.tX, pCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictCrdImage.hDC, 12 * mCardWith, 0, SRCCOPY)
End Sub


Sub SubSetMoveableToTablePile(pPlacePileInPile As Integer)
'Set the moveable bit to all piles in TablePile
Dim i As Integer
    For i = SubHowManyCardsInPile(pPlacePileInPile) To 1 Step True
        SubSetCardMoveable pPlacePileInPile, i, FuncGetMoveableFromTablePile(pPlacePileInPile, i)
    Next i
End Sub

Sub SubSetMoveableToFundationPile(ByVal pPile As Integer, ByVal pMoveable As Integer)
'Set the moveable bit to all piles in FundationPile
    mFundationPileInfo(pPile).tPileInfoMoveable = pMoveable
End Sub

Sub SubSetMoveableToDiscardPile(ByVal pPileInPile As Integer, ByVal pCardIndex As Integer, ByVal pMoveable As Integer)
'Set the moveable bit to pile in DiscardPile
    mDiscardPileInfo(pPileInPile, pCardIndex).tPileInfoMoveable = pMoveable
End Sub

Sub SubSetMoveableToDealPile(ByVal pPileInPile As Integer, ByVal pCardIndex As Integer, ByVal pMoveable As Integer)
'Set the moveable bit to pile in DealPile
    mDealPileInfo(pPileInPile, pCardIndex).tPileInfoMoveable = pMoveable
End Sub

Function FuncGetCardPicture() As Integer
    FuncGetCardPicture = mBackCardPicture
End Function

Function FuncGetAnimationSpeed() As Single
    FuncGetAnimationSpeed = mAnimationSpeed
End Function


Function CheckCardsPermissions(pCard As CardInfo, pCardDest As CardInfo, ByVal pRule As Integer) As Integer
    If pCard.tCardSuit = 0 Or pCardDest.tCardSuit = 0 Then
        CheckCardsPermissions = False
    Else
        Select Case pRule
        Case gFundationRule
            CheckCardsPermissions = (pCard.tCardSuit = pCardDest.tCardSuit) And (pCard.tCardValue = pCardDest.tCardValue + 1)
        Case gTableRule
            CheckCardsPermissions = (mSuit(pCard.tCardSuit) <> mSuit(pCardDest.tCardSuit)) And (pCard.tCardValue = pCardDest.tCardValue + 1)
        End Select
    End If
End Function


Sub SubGetInfoFwin(ByVal pPileInPile As Integer, ByVal pCardsInPile As Integer, pCard As CardInfo)
'sub need for fast win
    pCard = mTablePileInfo(pPileInPile, pCardsInPile).tPileInfoCard
End Sub

Function SubHowManyCardsInPile(ByVal pPlacePileInPile As Integer) As Integer
    SubHowManyCardsInPile = mCardSelectedInPile(pPlacePileInPile)
End Function


Function FuncGetMoveableFromTablePile(ByVal pPileInPile As Integer, ByVal pCardsInPile As Integer) As Integer
    FuncGetMoveableFromTablePile = mTablePileInfo(pPileInPile, pCardsInPile).tPileCardBlocked
End Function


Function FuncCpuGetInfoMoveableTablePile(ByVal pPileInPile As Integer, ByVal pCardsInPile As Integer) As Integer
    FuncCpuGetInfoMoveableTablePile = mTablePileInfo(pPileInPile, pCardsInPile).tPileInfoMoveable
End Function

Sub SubSetCardMoveable(ByVal pPileInPile As Integer, ByVal pCardsInPile As Integer, ByVal pMoveable As Integer)
    mTablePileInfo(pPileInPile, pCardsInPile).tPileInfoMoveable = pMoveable
End Sub

Sub SubTurnFaceUpTable(ByVal pPileInPileSource As Integer, ByVal pCardIndex As Integer, ByVal pMoveable As Integer, ByVal pUndo As Integer)
'turn face up in table pile
Dim lReturn As Long
Dim lRet As Integer

    If gSoundTurnedOn Then
        lReturn = sndPlaySound(gSoundTurn, SND_ASYNC) 'turn face up in table pile
    End If
    
    mTablePileInfo(pPileInPileSource, pCardIndex).tPileCardBlocked = pMoveable
    SubDrawCardInPile mTablePileInfo(pPileInPileSource, pCardIndex), False, mCard, 18
    
    If pUndo Then
        lRet = FuncSaveMove(TablePile, pPileInPileSource, TablePile, pPileInPileSource, pCardIndex)
    End If
    
End Sub


Sub SubGetPileInfo(ByVal pPilePlaceInPile As Integer, pPile As PILESINFORMATION)
Dim i As Integer
Dim lCardSelectedInPile As Integer

    pPile.tnPileNumber = TablePile
    pPile.tnPileNumberInPile = pPilePlaceInPile
    lCardSelectedInPile = mCardSelectedInPile(pPilePlaceInPile)
    
    If lCardSelectedInPile = 0 Then
        pPile.tCardIndex = 0
        pPile.tnPileInfo = mTablePileInfo(pPilePlaceInPile, 1)
        Exit Sub
    End If
    
    If Not mTablePileInfo(pPilePlaceInPile, lCardSelectedInPile).tPileInfoMoveable Then
        pPile.tCardIndex = 0
        pPile.tnPileInfo = mTablePileInfo(pPilePlaceInPile, lCardSelectedInPile)
        Exit Sub
    End If
    
    For i = lCardSelectedInPile - 1 To 1 Step True
        If Not mTablePileInfo(pPilePlaceInPile, i).tPileInfoMoveable Then
            pPile.tCardIndex = i + 1
            pPile.tnPileInfo = mTablePileInfo(pPilePlaceInPile, i + 1)
            Exit Sub
        End If
    Next i
    
    pPile.tCardIndex = 1
    pPile.tnPileInfo = mTablePileInfo(pPilePlaceInPile, 1)
End Sub


Sub SubInitializeDeck(pForm As Form, pAnimationInterval As Single, pAnimationSpeed As Single, pBmpFile As String)
Dim i As Integer
Dim j As Integer
    Set mForm = pForm
    Randomize
    
    mSuit(gClubs) = gBlackCard
    mSuit(gDiamonds) = gRedCard
    mSuit(gHearts) = gRedCard
    mSuit(gSpades) = gBlackCard
    
    mCard.tCardSuit = 0
    mCard.tCardValue = 0
    
    
    
    For i = 1 To 4 ' 4 suites
        For j = 1 To 13 '13 cards
            mDeck((i - 1) * 13 + j).tCardSuit = i
            mDeck((i - 1) * 13 + j).tCardValue = j
        Next j
    Next i
    
    
    mBackCardPicture = 2
    nCardsDroped = 0
    mAnimationSpeed = pAnimationSpeed ' speed of animation cards
    mAnimationInterval = pAnimationInterval 'interval for cards animation
    mUndosLeft = 1
    mTempUndos = 1
    SubLoadBmpFile pBmpFile
End Sub


Function SubCardMove(pPile As PILESINFORMATION, pPileDest As PILESINFORMATION, ByVal nCount As Integer, ByVal pUndo As Integer, ByVal pUpdateBackground As Integer) As Integer
Dim i As Integer
Dim lCardsTotalInPileAfterMove As Integer
Dim lCard() As CardInfo


    ReDim lCard(1 To nCount) As CardInfo

    If pPile.tnPileNumber = pPileDest.tnPileNumber And pPile.tnPileNumberInPile = pPileDest.tnPileNumberInPile Then
         Exit Function
     End If


    For i = 1 To nCount
        Select Case pPile.tnPileNumber
           Case TablePile
            SubGetCardFromTablePile pPile.tnPileNumberInPile, pUpdateBackground, lCard(i)
        Case FundationPile
        'When the player get a card in fundation pile(if is actived in menu)
            SubGetCardFromFundation pPile.tnPileNumberInPile, pUpdateBackground, lCard(nCount - i + 1)
        Case DiscardPile
        'When the player get a card in discard pile, refrecha a outra carta
            SubGetCardFromDiscardPile pPile.tnPileNumberInPile, pUpdateBackground, lCard(i)
        Case DealPile
            SubGetCardFromDealPile pPile.tnPileNumberInPile, pUpdateBackground, lCard(i)
        End Select
    Next i

    For i = nCount To 1 Step True
        Select Case pPileDest.tnPileNumber
        Case TablePile
            SubPutCardsInTablePile lCard(i), pPileDest.tnPileNumberInPile, gFaceUp, pUpdateBackground
            lCardsTotalInPileAfterMove = mCardSelectedInPile(pPileDest.tnPileNumberInPile)
        Case FundationPile
            SubPutCardsInFundationPile lCard(nCount - i + 1), pPileDest.tnPileNumberInPile, pUpdateBackground
        Case DiscardPile
            SubPutCardsInDiscardPile lCard(i), pPileDest.tnPileNumberInPile, gFaceUp, pUpdateBackground
        Case DealPile
            SubPutCardsInDealPile lCard(i), pPileDest.tnPileNumberInPile, gFaceUp, pUpdateBackground
        End Select
    Next i

    If pUndo Then
        SubCardMove = FuncSaveMove(pPile.tnPileNumber, pPile.tnPileNumberInPile, pPileDest.tnPileNumber, pPileDest.tnPileNumberInPile, nCount)
       Else
        SubCardMove = False
    End If

End Function

Private Function FuncCheckIfInsideFundations(pSource As tCordinates, pDest As tCordinates) As Integer
    FuncCheckIfInsideFundations = (0 <= pSource.tX - pDest.tX) And (pSource.tX - pDest.tX < mCardWith) And (0 < pSource.tY - pDest.tY) And (pSource.tY - pDest.tY < mCardHeight)
End Function

Private Sub SubWherePlayerHaveClicked(pCords As tCordinates, pPile As PILESINFORMATION)
Dim i As Integer
Dim j As Integer
Dim lCords As tCordinates

     'check if player have clicked in table pile
    For i = 1 To 7
        If mCardSelectedInPile(i) = 0 Then
            lCords = mTablePileInfo(i, 1).tPileInfoCords
            If FuncCheckIfInsideFundations(pCords, lCords) Then
                pPile.tnPileNumber = TablePile
                pPile.tnPileNumberInPile = i
                pPile.tCardIndex = 0
                pPile.tnPileInfo = mTablePileInfo(i, 1)
                Exit Sub
            End If
        Else
            For j = mCardSelectedInPile(i) To 1 Step True
                lCords = mTablePileInfo(i, j).tPileInfoCords
                If FuncCheckIfInsideFundations(pCords, lCords) Then
                    pPile.tnPileNumber = TablePile
                    pPile.tnPileNumberInPile = i
                    pPile.tCardIndex = j 'return the card index in pile
                    pPile.tnPileInfo = mTablePileInfo(i, j)
                    Exit Sub
                End If
            Next j
        End If
    Next i
     
     'check if player have clicked in fundation pile
    For i = 1 To 4
        lCords = mFundationPileInfo(i).tPileInfoCords
        If FuncCheckIfInsideFundations(pCords, lCords) Then
        'player have clicked in fundation pile
            pPile.tnPileNumber = FundationPile
            pPile.tnPileNumberInPile = i
            If mFundationPileInfo(i).tPileInfoCard.tCardSuit = 0 Then
                pPile.tCardIndex = 0 'the fundation pile is empty
            Else
                pPile.tCardIndex = 1 'the fundation pile have 1 or more cards
            End If
            pPile.tnPileInfo = mFundationPileInfo(i)
            Exit Sub
        End If
    Next i
    
     'check if player have clicked in discard pile
        If mDiscardPileCards(1) = 0 Then
            lCords = mDiscardPileInfo(1, 1).tPileInfoCords
            If FuncCheckIfInsideFundations(pCords, lCords) Then 'the discard pile as no cards
                pPile.tnPileNumber = DiscardPile
                pPile.tnPileNumberInPile = 1
                pPile.tCardIndex = 0 'return no cards in discard pile
                pPile.tnPileInfo = mDiscardPileInfo(1, 1)
                Exit Sub
            End If
        Else
            For j = mDiscardPileCards(1) To 1 Step True
                lCords = mDiscardPileInfo(1, j).tPileInfoCords
                If FuncCheckIfInsideFundations(pCords, lCords) Then
                    'player have clicked or droped a card in discard pile
                    pPile.tnPileNumber = DiscardPile
                    pPile.tnPileNumberInPile = 1
                    pPile.tCardIndex = j
                    pPile.tnPileInfo = mDiscardPileInfo(1, j)
                    Exit Sub
                End If
            Next j
        End If
    
    
    'check if player have clicked in deal pile
        If mHowManyCardsLeft(1) = 0 Then
            lCords = mDealPileInfo(1, 1).tPileInfoCords
            If FuncCheckIfInsideFundations(pCords, lCords) Then 'the deal pile is empty
                pPile.tnPileNumber = DealPile
                pPile.tnPileNumberInPile = 1
                pPile.tCardIndex = 0 'return no cards in deal pile
                pPile.tnPileInfo = mDealPileInfo(1, 1)
                Exit Sub
            End If
        Else
            For j = mHowManyCardsLeft(1) To 1 Step True
                lCords = mDealPileInfo(1, j).tPileInfoCords
                If FuncCheckIfInsideFundations(pCords, lCords) Then
                    'player have clicked in deal pile
                    pPile.tnPileNumber = DealPile
                    pPile.tnPileNumberInPile = 1
                    pPile.tCardIndex = j 'return how many cards left in deal pile
                    pPile.tnPileInfo = mDealPileInfo(1, j)
                    Exit Sub
                End If
            Next j
        End If
    
    
    pPile.tnPileNumber = 0 'the player have clicked in green table
End Sub

Sub SubPutImgCardsInBuffer(pForm As Form)
Dim i As Integer
Dim lReturn As Integer
    pForm.Cls
    
    If gSoundTurnedOn Then 'play the surfle sound
        lReturn = sndPlaySound(gSoundShuffle, SND_ASYNC)
    End If
    'desenha as 7 fundacões para o table pile
    For i = 1 To 7 'for table piles
        mCardSelectedInPile(i) = 0
        mTablePileInfo(i, 1).tPileInfoCard = mCard
        SubDrawTablePlaces mTablePileInfo(i, 1).tPileInfoCords
    Next i
    
    'desenha as 4 fundações para as fundations piles
    For i = 1 To 4 'for fundation piles
        mFundationPileInfo(i).tPileInfoCard = mCard
        SubDrawFundationPlaces mFundationPileInfo(i).tPileInfoCords
    Next i
    
    'desenha 1 fundaçõa para discard pile
            mDiscardPileCards(1) = 0
            mDiscardPileInfo(1, 1).tPileInfoCard = mCard
            SubDrawDiscardPlaces mDiscardPileInfo(1, 1).tPileInfoCords
    
    'desenha uma fundação para o deal pile
            mHowManyCardsLeft(1) = 0
            SubDrawDealPlaces mDealPileInfo(1, 1).tPileInfoCords
    
    SubSurfleDeck
    mUndosLeft = 1
    mTempUndos = 1
End Sub

Sub SubAddCardToDeck(pCard As CardInfo)
    If mTotalCards = gTotalCardsInDeck Then
        pCard = mCard 'all the cards in deck as deal
    Else
        mTotalCards = mTotalCards + 1
        pCard = mDeck(mTotalCards)
    End If
End Sub

Sub SubRefreshPilesBackground(ByVal pClear As Integer)
Dim i As Integer
    If pClear Then
        mForm.Cls
    End If
    
    For i = 1 To 7
        SubRefreshAllPilesBackground TablePile, i, pClear
    Next i
    
    For i = 1 To 4
        SubRefreshAllPilesBackground FundationPile, i, pClear
    Next i
    
    SubRefreshAllPilesBackground DiscardPile, 1, pClear
    
    SubRefreshAllPilesBackground DealPile, 1, pClear
    
End Sub

Private Sub SubRefreshAllPilesBackground(ByVal pPile As Integer, ByVal pPileNumberInPile As Integer, ByVal pBackground As Integer)
Dim i As Integer

    Select Case pPile
    Case TablePile
        If mCardSelectedInPile(pPileNumberInPile) = 0 Or pBackground Then
            SubDrawTablePlaces mTablePileInfo(pPileNumberInPile, 1).tPileInfoCords
        End If
        If mCardSelectedInPile(pPileNumberInPile) <> 0 Then
            For i = 1 To mCardSelectedInPile(pPileNumberInPile)
                If i = 1 Then
                    SubDrawCardInPile mTablePileInfo(pPileNumberInPile, i), pBackground, mCard, 18
                Else
                    SubDrawCardInPile mTablePileInfo(pPileNumberInPile, i), pBackground, mTablePileInfo(pPileNumberInPile, i - 1).tPileInfoCard, 18
                End If
            Next i
        End If
        
    Case FundationPile
        If mFundationPileInfo(pPileNumberInPile).tPileInfoCard.tCardSuit = 0 Or pBackground Then
            SubDrawFundationPlaces mFundationPileInfo(pPileNumberInPile).tPileInfoCords
        End If
        If mFundationPileInfo(pPileNumberInPile).tPileInfoCard.tCardSuit <> 0 Then
            SubDrawCardInPile mFundationPileInfo(pPileNumberInPile), pBackground, mCard, 0
        End If
        
    Case DiscardPile
        If mDiscardPileCards(pPileNumberInPile) = 0 Or pBackground Then
            SubDrawDiscardPlaces mDiscardPileInfo(pPileNumberInPile, 1).tPileInfoCords
        End If
        'when player click in menu Refresh this draw the first card in pile
            For i = 1 To mDiscardPileCards(pPileNumberInPile)
                If i = 1 Then
                    SubDrawCardInPile mDiscardPileInfo(pPileNumberInPile, i), pBackground, mCard, 1
                Else
                    SubDrawCardInPile mDiscardPileInfo(pPileNumberInPile, i), pBackground, mDiscardPileInfo(pPileNumberInPile, i - 1).tPileInfoCard, 1
                End If
            Next i
        
    Case DealPile
        If mHowManyCardsLeft(pPileNumberInPile) = 0 Or pBackground Then
            SubDrawDealPlaces mDealPileInfo(pPileNumberInPile, 1).tPileInfoCords
        End If
        If mHowManyCardsLeft(pPileNumberInPile) <> 0 Then
            For i = 1 To mHowManyCardsLeft(pPileNumberInPile)
                If i = 1 Then
                    SubDrawCardInPile mDealPileInfo(pPileNumberInPile, i), pBackground, mCard, 0
                Else
                    SubDrawCardInPile mDealPileInfo(pPileNumberInPile, i), pBackground, mDealPileInfo(pPileNumberInPile, i - 1).tPileInfoCard, 0
                End If
            Next i
        End If
    End Select
    
End Sub

Private Sub SubGetCardFromDealPile(ByVal pPileNumberInPile As Integer, ByVal pUpdateBackground As Integer, pCard As CardInfo)
Dim lTotalCardsInPile As Integer
Dim lRet As Integer

    lTotalCardsInPile = mHowManyCardsLeft(pPileNumberInPile)
    pCard = mDealPileInfo(pPileNumberInPile, lTotalCardsInPile).tPileInfoCard
    If pUpdateBackground Then
    'em modo automatico, quando sai uma carta do dealpile directamente para as fundationpile
        lRet = bitblt(mForm.hDC, mDealPileInfo(pPileNumberInPile, lTotalCardsInPile).tPileInfoCords.tX, mDealPileInfo(pPileNumberInPile, lTotalCardsInPile).tPileInfoCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictBackground.hDC, (pCard.tCardValue - 1) * mCardWith, pCard.tCardSuit * mCardHeight, SRCCOPY)
    End If
    
    mHowManyCardsLeft(pPileNumberInPile) = mHowManyCardsLeft(pPileNumberInPile) - 1
    mDealPileInfo(pPileNumberInPile, lTotalCardsInPile).tPileInfoCard = mCard
End Sub

Private Sub SubGetCardFromDiscardPile(ByVal pPileNumberInPile As Integer, ByVal pToDealPile As Integer, pCard As CardInfo)
'passa todas as cartas do discard para o deal pile
Dim lTotalCardsInDiscardPile As Integer
Dim lReturn As Integer
Dim lCardCords As tCordinates
    lTotalCardsInDiscardPile = mDiscardPileCards(pPileNumberInPile)
    pCard = mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile).tPileInfoCard
    
    If pToDealPile Then 'when card is move to deal pile, remove last card
        lCardCords = mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile).tPileInfoCords
        lReturn = bitblt(mForm.hDC, lCardCords.tX, lCardCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictBackground.hDC, (pCard.tCardValue - 1) * mCardWith, pCard.tCardSuit * mCardHeight, SRCCOPY)
    End If
    
    mDiscardPileCards(pPileNumberInPile) = mDiscardPileCards(pPileNumberInPile) - 1
    mDiscardPileInfo(pPileNumberInPile, lTotalCardsInDiscardPile).tPileInfoCard = mCard
End Sub

Private Sub SubGetCardFromFundation(ByVal pPileNumberInPile As Integer, ByVal pUpdateBackground As Integer, pCard As CardInfo)
  'fundation pile
    pCard = mFundationPileInfo(pPileNumberInPile).tPileInfoCard
    If pCard.tCardValue = 1 Then 'if is a ACE
        mFundationPileInfo(pPileNumberInPile).tPileInfoCard = mCard 'no more cards in that pile
        If pUpdateBackground Then
            SubDrawFundationPlaces mFundationPileInfo(pPileNumberInPile).tPileInfoCords
        End If
    Else
    'set the card value to -1 card in pile
        mFundationPileInfo(pPileNumberInPile).tPileInfoCard.tCardValue = pCard.tCardValue - 1
        If pUpdateBackground Then
            SubDrawCardInPile mFundationPileInfo(pPileNumberInPile), False, mCard, 0
        End If
    End If
End Sub

Private Sub SubGetCardFromTablePile(ByVal pPileInPile As Integer, ByVal pUpdateBackground As Integer, pCard As CardInfo)
Dim lCardsInPile As Integer
Dim lRet As Integer
'table pile
    lCardsInPile = mCardSelectedInPile(pPileInPile)
    pCard = mTablePileInfo(pPileInPile, lCardsInPile).tPileInfoCard
    If pUpdateBackground Then
        lRet = bitblt(mForm.hDC, mTablePileInfo(pPileInPile, lCardsInPile).tPileInfoCords.tX, mTablePileInfo(pPileInPile, lCardsInPile).tPileInfoCords.tY, mCardWith, mCardHeight, frmImgCrdBuffers.pictBackground.hDC, (pCard.tCardValue - 1) * mCardWith, pCard.tCardSuit * mCardHeight, SRCCOPY)
    End If
    mCardSelectedInPile(pPileInPile) = mCardSelectedInPile(pPileInPile) - 1
    mTablePileInfo(pPileInPile, lCardsInPile).tPileInfoCard = mCard
End Sub




Sub SubSetCardBack(pCardBack As Integer)
'max is 6 pictures for card design
    If 1 <= pCardBack And pCardBack <= 6 Then
        mBackCardPicture = pCardBack
    End If
End Sub

Sub SubLoadBmpFile(pBmpFile As String)
    frmImgCrdBuffers.pictCrdImage = LoadPicture(pBmpFile)
    mCardHeight = frmImgCrdBuffers.pictCrdImage.ScaleHeight / 5
    mCardWith = frmImgCrdBuffers.pictCrdImage.ScaleWidth / 13
    frmImgCrdBuffers.pictBackground.Height = frmImgCrdBuffers.pictCrdImage.Height
    frmImgCrdBuffers.pictBackground.Width = frmImgCrdBuffers.pictCrdImage.Width
End Sub


Sub FuncSetAnimationSpeed(pSpeed As Single)
    mAnimationSpeed = pSpeed
End Sub


Sub SubInitPiles(pTablePile() As tCordinates, pFundationPile() As tCordinates, plDiscardPilePile() As tCordinates, plDealPilePile() As tCordinates)

Dim i As Integer
Dim j As Integer
    ReDim mTablePileInfo(1 To 7, 20) ' MAX of cards in Any TablePile -> 13 cards + 7 cards in pile = 20
    ReDim mCardSelectedInPile(1 To 7) 'Total of deal cards to put in 7 pile of TablePile
    
    For i = 1 To 7
        mTablePileInfo(i, 1).tPileInfoCords = pTablePile(i)
        For j = 1 To 7
            mTablePileInfo(i, j).tPileInfoCords.tX = mTablePileInfo(i, 1).tPileInfoCords.tX
            mTablePileInfo(i, j).tPileInfoCords.tY = mTablePileInfo(i, 1).tPileInfoCords.tY + (j - 1) * 18
        Next j
    Next i
    
    ReDim mFundationPileInfo(4)
    For j = 1 To 4
        mFundationPileInfo(j).tPileInfoCords = pFundationPile(j)
    Next j
    
        ReDim mDiscardPileInfo(1, 26) 'max cards in DiscardPile is 26
        ReDim mDiscardPileCards(1)
            mDiscardPileInfo(1, 1).tPileInfoCords = plDiscardPilePile(1)
            For j = 1 To 26
                mDiscardPileInfo(1, j).tPileInfoCords.tX = mDiscardPileInfo(1, 1).tPileInfoCords.tX
                mDiscardPileInfo(1, j).tPileInfoCords.tY = mDiscardPileInfo(1, 1).tPileInfoCords.tY + (j - 1)
            Next j
    
        ReDim mDealPileInfo(1, 26) 'max cards in DealPile is 26
        ReDim mHowManyCardsLeft(1)
            mDealPileInfo(1, 1).tPileInfoCords = plDealPilePile(1)
            For j = 1 To 26
                mDealPileInfo(1, j).tPileInfoCords.tX = mDealPileInfo(1, 1).tPileInfoCords.tX
                mDealPileInfo(1, j).tPileInfoCords.tY = mDealPileInfo(1, 1).tPileInfoCords.tY + (j - 1)
            Next j
End Sub

Sub SubSetSound(ByVal pOnOff As Integer)
    gSoundTurnedOn = pOnOff
End Sub

Private Sub SubSurfleDeck()
'Surfle the deck
Dim lCard As CardInfo
Dim i, j As Integer
Dim lRandom As Integer
    mTotalCards = 0
 For j = 1 To 10 'surfle 10 times
    For i = 1 To gTotalCardsInDeck
        lRandom = Int(Rnd * (gTotalCardsInDeck - 1)) + 1
        lCard = mDeck(lRandom)
        mDeck(lRandom) = mDeck(i)
        mDeck(i) = lCard
    Next
Next j
End Sub



Sub SubUndoMove()
Dim lUndos As UNDOTYPE
Dim pPile As PILESINFORMATION
Dim lPilesDest As PILESINFORMATION
Dim lRet As Integer
Const lNotMoveable = False

    If mUndosLeft = mTempUndos Then
        MsgBox "No more moves to undo, sorry.", MB_ICONASTERISK
        Exit Sub
    ElseIf mUndosLeft = 1 Then
        mUndosLeft = mMaxUndosPermited
    Else
        mUndosLeft = mUndosLeft - 1 ' number of undos left -1
    End If
    
    lUndos = mUndoBuffer(mUndosLeft)
    If lUndos.tUndoPileSource = TablePile And lUndos.tUndoPileDest = TablePile And lUndos.tUndoPileInPileDest = lUndos.tUndoPileInPileSource Then
     'turn faceup in TablePile
        SubTurnFaceUpTable lUndos.tUndoPileInPileSource, lUndos.tUndoTotalCardsMoved, lNotMoveable, gNoSaveUndo
    ElseIf lUndos.tUndoPileSource = DealPile And lUndos.tUndoPileDest = DealPile And lUndos.tUndoPileInPileDest = lUndos.tUndoPileInPileSource Then
     'turn faceup in DealPile
        SubTurnFaceUpDeal lUndos.tUndoPileInPileSource, lUndos.tUndoTotalCardsMoved, lNotMoveable, gNoSaveUndo
    Else
        pPile.tnPileNumber = lUndos.tUndoPileSource
        pPile.tnPileNumberInPile = lUndos.tUndoPileInPileSource
        lPilesDest.tnPileNumber = lUndos.tUndoPileDest
        lPilesDest.tnPileNumberInPile = lUndos.tUndoPileInPileDest
        lRet = SubCardMove(lPilesDest, pPile, lUndos.tUndoTotalCardsMoved, gNoSaveUndo, True)
    End If
End Sub

Function FuncGetScrunchValue(pPile As Integer) As Integer
    Select Case pPile
    Case TablePile
        FuncGetScrunchValue = 18
    Case Else
        FuncGetScrunchValue = 0
    End Select
End Function


Sub SubGetCardInfoFromTablePile(ByVal pPileNumberInPile As Integer, pPile As PILESINFORMATION)
    pPile.tnPileNumber = TablePile
    pPile.tnPileNumberInPile = pPileNumberInPile
    pPile.tCardIndex = mCardSelectedInPile(pPileNumberInPile) 'return card index in pile
    
    If mCardSelectedInPile(pPileNumberInPile) = 0 Then
        pPile.tnPileInfo = mTablePileInfo(pPileNumberInPile, 1)
    Else
        pPile.tnPileInfo = mTablePileInfo(pPileNumberInPile, mCardSelectedInPile(pPileNumberInPile))
    End If
    
End Sub

Sub SubGetCardInfoFromDiscardPile(ByVal pPileNumberInPile As Integer, pPile As PILESINFORMATION)
    pPile.tnPileNumber = DiscardPile
    pPile.tnPileNumberInPile = pPileNumberInPile
    pPile.tCardIndex = mDiscardPileCards(pPileNumberInPile)
    
    If mDiscardPileCards(pPileNumberInPile) = 0 Then
        pPile.tnPileInfo = mDiscardPileInfo(pPileNumberInPile, 1)
    Else
        pPile.tnPileInfo = mDiscardPileInfo(pPileNumberInPile, mDiscardPileCards(pPileNumberInPile))
    End If
End Sub

Sub SubGetCardInfoFromDealPile(ByVal pPileNumberInPile As Integer, pPile As PILESINFORMATION)
    pPile.tnPileNumber = DealPile ' the number of deal pile -> 4
    pPile.tnPileNumberInPile = pPileNumberInPile
    pPile.tCardIndex = mHowManyCardsLeft(pPileNumberInPile) 'return how many cards left in pile
    
    If mHowManyCardsLeft(pPileNumberInPile) = 0 Then
        pPile.tnPileInfo = mDealPileInfo(pPileNumberInPile, 1) ' if deal pile no cards
    Else
        pPile.tnPileInfo = mDealPileInfo(pPileNumberInPile, mHowManyCardsLeft(pPileNumberInPile))
    End If
End Sub

Sub SubGetCardInfoFromFundationPile(ByVal pPileNumberInPile As Integer, pPile As PILESINFORMATION)
    pPile.tnPileNumber = FundationPile
    pPile.tnPileNumberInPile = pPileNumberInPile
    If mFundationPileInfo(pPileNumberInPile).tPileInfoCard.tCardSuit = 0 Then
        pPile.tCardIndex = 0
    Else
        pPile.tCardIndex = 1
    End If
    pPile.tnPileInfo = mFundationPileInfo(pPileNumberInPile)
End Sub

