## Message/UserForm Design notes
The message form is organized in a hierarchy of frames as follows.

    +----Message Area (Frame)----------------+
    | +---Message Section (Frame)----------+ |
    | | Message Section Label (Label)      | |
    | | +--Message Section Text (Frame)--+ | |
    | | | Message Section (TextBox)      | | |
    | | +--------------------------------+ | |
    | +------------------------------------+ |
    +----------------------------------------+
    +----Reply Area (Frame)------------------+
    | +----Replies Row (Frame)-------------+ |
    | | Replies (CommandButtons)           | |
    | +------------------------------------+ |
    +----------------------------------------+
  * The controls are not addressed through their object name but via collections build at the UserForm's initialization by using the parent property.
  * The number of available message sections and reply CommandButtons is exclusively specified through the UserForm's design - the code can handle any number of it without change.  
  As an example: In case the last (third) message section is duplicated, four instead of just three sections are available - regardless of being used.


    Private Sub CollectControls()
        Collect into:=cllAreas, ctltype:="Frame",  fromparent:=Me

        ' Collect Message Sections
        Collect into:=cllMsgSections, ctltype:="Frame", fromparent:=cllAreas(1)
        ' Collect Message Section Labels
        Collect into:=cllMsgSectionLabels, fromparent:=cllMsgSections, ctltype:="Label"
        ' Collect Message Section Text Frames
        Collect into:=cllMsgSectionTextFrame, ctltype:="Frame", fromparent:=cllMsgSections
        ' Collect Message Section TextBoxes
        Collect into:=cllMsgSectionText, ctltype:="TextBox", fromparent:=cllMsgSectionsTextFrame

        ' Collect Reply Rows
        Collect into:=cllReplyRows, ctltype:="Frame", fromparent:=cllAreas(2)
        ' Collect for each Reply Row the Replies
        For each v in cllReplyRows        
            Collect into:=cllRepliesRow,  ctltype:="CommandButton", fromparent:=v
             cllRepliesRows.Add cllRepliesRow       
        Next v

    End Sub

    Private Sub Collect(ByRef into As Collection, _
               ByVal fromparent As Object, _
               ByVal ctltype As String)

        Dim ctl As MsForms.Control    
         
        Set into = Nothing: Set into = New Collection
        Select Case TypeName(fromparent)
            Case "Frame", "UserForm"
                For each ctl in Me
                    If TypeName(ctl) = ctltype And ctl.Parent Is fromparent _
                    Then into.Add ctl
                Next ctl
            Case "Collection"
                For each v in fromparent
                    For each ctl in Me
                        If TypeName(ctl) = ctltype And ctl.Parent Is v _
                        Then into.Add ctl
                   Next ctl
                Next v
        End Select

    End Sub

## Width Adjustment
The message form is initialized with the specified minimum message form width. Width expansion may be  triggered by the setup (in the outlined sequence) of the following width determining elements:
  1. **Title**  
When the **Title** exceeds the specified  maximum message form width some text will be truncated. However, with a default maximum message form width of 80 % of the screen width that will happen pretty unlikely.

  2. **Mono-spaced message section** followed by **Replies Rows**  
When either of the two exceeds the maximum message form width it will get a horizontal scroll bar.

  3. **Proportional spaced message sections**  
are setup at last and adjusted to the (by then) final message form width.



    ' Re-adjust width of message section text and
    ' adjust frames height accordingly
    ' ---------------------------------------------
    Private Sub MsgSectionAdjustHeightToAvailableWidth( _
            ByVal section As Long, _
            ByVal newwidth As Single)

        Dim s As String
        Dim siNewHeight As Single
     
        With MsgSectionText(section)
            s = .Value
            .Value = vbNullstring
             .AutoSize = False
            .Width = newwidth
            .MultiLine = True
            .AutoSize = True
            .Value = s
            MsgSectionTextFrame.Height = .Height + F_MARGIN
            MsgSectionTextFrame.Width = .Width + F_MARGIN
        End With
        
    End Sub

## Height Adjustment
### Height Increment
Height size increment is done along with
- the setup of a message section by the subsequent repositioning of all below displayed elements' top position
- the setup of the reply buttons (reply button rows respectively).

These height increments are done without considering  the specified maximum message form height.

### Height Decrement
When all elements are setup and the message form exceeds the maximum specified height the form height the message area and/or the reply area are adjusted. When the areas' height relation is 50/50 to 65/35 both areas will get a vertical scroll bar and the height is decremented by the corresponding relation. Otherwise only the taller area is reduced by the exceeding amount and gets a vertical scroll bar. The width of the scrollbar is the height before the reduction 

    '   
    Private Function MsgAreaHeight() As Single
    
    End Function


## Vertical Repositioning
Adjusting the top position of displayed elements is due initially when an element had need setup and subsequently whenever an element's height changed because of a width adjustment. Together with the adjustment of the top position of the bottommost element the new height of the message form is set.

Note: This top repositioning may be done just once when all elements had initially been  setup. However, for testing it is more appropriate to be performed immediately after setup of each individual element.

    Private Sub RepositionTop()
        ReposTopMsgSections
        ReposTopPosReplyRows
        ReposTopAreas
    End Sub




