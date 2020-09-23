VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SatAVL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'_______________________________________________________________________________
'
'ABSTRACT
'   Provides any amount of requested memory (as long as there is available RAM)
'   and associates that memory with a user-defined alphanumeric key. Features
'   very fast retrieval in sorted order, even faster sequential access (in
'   the chronological order the requests were made).
'_______________________________________________________________________________
'
'DESCRIPTION
'   The alphanumeric keys are associated with a certain quantity of bytes located
'   at a designated position in RAM. That RAM is and remains exclusively reserved
'   for that Key. Via the key the address can be returned, and also the reserved
'   bytes at that address. Keys are internally managed with a balanced binary AVL
'   tree. It is possible to iterate through the whole collection of keys in the
'   order in which these keys were created (and the memory allocated), returning
'   the Key, the Address and the Size. It is also possible to access all the
'   Keys in sorted order (ascending as well as descending). All operations are
'   considerabily faster than VB's Collection object (as far as provided by the
'   Collection).
'_______________________________________________________________________________
'
'USAGE
'   What you do with the provided memory is up to you; however this is optimal
'   to store objects' addresses (available via ObjPtr) to circumvent the slow
'   COM mechanisms.
'_______________________________________________________________________________
'
'WARNINGS
'   * Whatever you do with the allocated memory, do not forget to free it
'     ("Set ThisClassesObject = Nothing" will do), or you soon end up with a lot
'     of blocked RAM which won't be freed until a restart.
'   * Do not attempt to write data beyond the end of the requested and allocated
'     memory fragment! Chances are you get a crash if you do.
'_______________________________________________________________________________
'
'AUTHOR
'   This software is (c) 2005 by Herbert Glarner.
'   The used TypeLib for the IMalloc interface is (c) by Brad Martinez.
'_______________________________________________________________________________
'
'HISTORY
'   04 Apr 2005    1.0          Creation and Access functionality
'   05 Apr 2005    1.1          Replaced StrComp, Keys stored as Byte Arrays
'   07 Apr 2005                 Fixed a bug in IStrComp
'_______________________________________________________________________________
'
'PROJECT FILES
'   SatAVL.cls                  (This file)
'_______________________________________________________________________________
'
'USED TEST FORM
'   SatAVLTest.frm (Test scenarioes and parallel comparison with VB Collection)
'_______________________________________________________________________________
'
'REQUIRED TYPE LIBRARIES
'   ISHF_Ex.tlb                 "IShellFolder Extended type Library v1.2"
'                               (The TypeLib is (c) by Brad Martinez)
'_______________________________________________________________________________
'
'REQUIRED DLLS
'   shell32                     (for "SHGetMalloc")
'_______________________________________________________________________________
'
'REQUIRED CONTROLS
'   -
'_______________________________________________________________________________
'
'EVENTS
'   -
'
'PUBLIC PROPERTIES
'   R           Nodes =     Count               Use to traverse by index 0...n-1
'
'PUBLIC METHODS
'   Address =   Add (Key, Bytes)                already existing: 0, else RAM
'
'LEXICAL ACCESS (SORTED BY KEY)
'   Index =     Lowest/Higher                   ascending, No more items: -1
'   Index =     Highest/Lower                   descending, No more items: -1
'
'DIRECT ACCESS (with result, use sequential access methods "Address" etc.)
'   Index =     Item (Key)                      not existing returns -1
'
'SEQUENTIAL ACCESS (IN APPENDING ORDER), ALSO DIRECT ACCESS WITH KNOWN INDEX
'   Key =       Key (Index)                     Index=0...n-1, n.ex. returns ""
'   Address =   Address (Index)                 Index=0...n-1, n.ex. returns 0
'   Bytes =     Size (Index)                    Index=0...n-1, n.ex. returns 0
'_______________________________________________________________________________
'
'EXTERNAL DLL METHODS

'IMalloc is available via the type library "ISHF_Ex.tlb" which has to be
'included into the project's references when used in the IDE.
Private Const NOERROR = 0
Private Declare Function SHGetMalloc Lib "shell32" _
    (ppMalloc As IMalloc) As Long
'_______________________________________________________________________________
'
'ERRORS (Base Number for this class: 4k)
'(Dont't forget to initialize descriptions in the Constructor)
Private Const ErrClass As String = "SatAVL"
Private Enum ErrNumber
    ErrBase = 4& * 1024&                            'Raised in:
    ErrNoMallocInterface = ErrBase + 0              'Initialize
    ErrFatalAdd = ErrBase + 1                       'Add
    ErrFatalItem = ErrBase + 2                      'Item
    'Always leave as the last entry:
    ErrAutoMax
End Enum
'Descriptions initialized on Construction
Private ErrDesc(ErrBase + 0 To ErrAutoMax - 1) As String
'_______________________________________________________________________________
'
'CONSTANTS

'Increasing elements vPath() if neccessary.
Private Const NEWELEMENTS As Long = 10

'Elements in rNode having no pointer to another index.
Private Const sstNoChild As Long = -1
Private Const sstNoParent As Long = -1

'StrComp results.
Private Const sstKeyIsSmaller As Long = -1
Private Const sstKeyIsEqual As Long = 0
Private Const sstKeyIsLarger As Long = 1

'Used in Balance calculation for Height fields of the UDT "rChunk".
Private Const ssmbLeftHeavy As Integer = -2
Private Const ssmbLeftBalanced As Integer = -1
Private Const ssmbBalanced As Integer = 0
Private Const ssmbRightBalanced As Integer = 1
Private Const ssmbRightHeavy As Integer = 2
'_______________________________________________________________________________
'
'ENUMS

'Used in the Direction field of the UDT rPath (see below).
Private Enum ssmDirection
    ssmdLeft
    ssmdRight
End Enum
'_______________________________________________________________________________
'
'UDTS

'Able to handle up to 2G nodes (signed Long referring to indices).
Private Type rNode
    'The key and a reference to it's content. (I've tested Key being an
    'Integer() array, converted into a such from an ordinary key via mapping
    'in Add(), so that only new keys needed to be mapped and the existing ones
    'were in Integer() format. However, key comparison takes up to 50% longer
    'time than the already slow string comparison, so it's strings again.)
    
    '(Seems that I could manage a slightly faster string comparison with
    'Byte arrays.)
    ByteKey() As Byte
    'Key As String                       'Provided by user, must be unique
    
    Address As Long                     'Memory allocated via Malloc
    Bytes As Long                       'As per user request, any number
    'Some links regarding the balanced binary tree ("Index" is the index of an
    'array in which this rChunk record is held among all relatives). Here we
    'definitely need Longs to care for more than 2^15 nodes.
    LeftTree As Long                    'Index of left subtree
    RightTree As Long                   'Index of right subtree
    Parent As Long                      'Index of the parent entry
    '(Integers rather Longs here to save 4 bytes; Integers with 2^15 entries for
    'the height should do the job: never encontered something such high due to
    'tree balancing. Approx. 1 million nodes need height < 40. Still I feel
    'that a byte is a risk, and there's also no need to save 2 bytes, since
    'alignment will use them anyway. I've also tested Long'anyway with writing and
    'accessing 1 million nodes: The time difference is not noticeable (once more,
    'once less time).
    LeftHeight As Integer               'This node's left height (excl. this node)
    RightHeight As Integer              'This node's right height (excl. this node)
End Type

'Detailled description see VARIABLES for "vPath()".
Private Type rPath
    Index As Long
    Direction As ssmDirection           'Left or right Path
End Type
'_______________________________________________________________________________
'
'EVENTS
'   None
'_______________________________________________________________________________
'
'VARIABLES

'Holds a reference to the IMalloc Interface, available via the type library
'ISHF_Ex.tlb to which a VB Project reference must be set to use it. Set on
'object construction, and destroyed on object destruction. The pointer is not
'made available to the external world; it is used object internally only to
'allocate resp. free memory.
Private ifMalloc As IMalloc

'This vector holds the balanced binary tree. In case that no "Alloc" was made
'so far, it is completely empty. With the first "Alloc" a first element (the
'root) is allocated.
Private vBinTree() As rNode
Private lBinTreeMax As Long     'The same as UBound(vBinTree)
Private lBinTreeNext As Long    '0=empty, 1=only 0 exists etc.
Private lBinTreeRoot As Long    'Initially 0, later anywhere (due to balancing)

'When adding a node, we need to keep track of the visited path, because we do
'not know before the very insertion if a tree grows higher or not. If it does,
'this may affect some or all nodes back up to the root, which we do not want to
'track again when we already have visted them: Instead their indices and also
'the direction we took (left/right) is held in this vector.
Private vPath() As rPath        'Never shrinks
Private lPathMax As Long        'The same as UBound(vPath)
Private lPathAct As Long        'Current pointer (0=start)

'Last accessed node via "Item", store the continuation points for some public
'methods like "Higher", "Lower" etc.
Private bLastAcc() As Byte      'Last searched Key via "Item".
Private lLastAcc As Long        'It's index; any "Add" invalidates this (-1).
'_______________________________________________________________________________



'CONSTRUCTION AND DESTRUCTION
'
Private Sub Class_Initialize()
    'Load error messages
    ErrDesc(ErrNoMallocInterface) = "Creation of a reference to the IMalloc interface failed."
    ErrDesc(ErrFatalAdd) = "Fatal Error in Add"
    ErrDesc(ErrFatalItem) = "Fatal Error in Item"

    'Returns a reference to the IMalloc interface. With it we can create and
    'destroy memory. Destroyed on object destruction.
    If Not (SHGetMalloc(ifMalloc) = NOERROR) Then
        'We could not obtain a refernce to the interface.
        Err.Raise ErrNoMallocInterface, ErrClass, ErrDesc(ErrNoMallocInterface)
    End If

    'Create some initial nodes for the binary tree. Note, that we are going to
    '*double* this as soon as it does not suffice anymore (adding a constant
    'number *always* becomes a bottleneck at some point of time (when adding lots
    'of elements), even if it's 10000).
    lBinTreeMax = 2&
    ReDim vBinTree(0& To lBinTreeMax) As rNode
    lBinTreeNext = 0&
    lBinTreeRoot = 0&

    'Initial vPath() dimensioning to avoid testing for an empty vector.
    lPathMax = NEWELEMENTS - 1&
    ReDim vPath(0& To lPathMax) As rPath
End Sub

Private Sub Class_Terminate()
    'Walking the whole binary tree and freeing all allocated memory.
    If vBinTree(lBinTreeRoot).Address > 0 Then
        FreeMemory lBinTreeRoot
    End If
    Erase vBinTree
    
    'Destroy the reference to the IMalloc object, which was created on object
    'Construction.
    Set ifMalloc = Nothing
End Sub

'Recursive procedure to free all allocated memory on object destruction.
Private Sub FreeMemory(ParentIndex As Long)
    Dim lChildIndex As Long
    
    lChildIndex = vBinTree(ParentIndex).LeftTree
    If lChildIndex <> sstNoChild Then FreeMemory lChildIndex
    lChildIndex = vBinTree(ParentIndex).RightTree
    If lChildIndex <> sstNoChild Then FreeMemory lChildIndex
    
    'Freeing the memory for this ParentIndex
    ifMalloc.Free ByVal vBinTree(ParentIndex).Address
End Sub
'_______________________________________________________________________________



'PUBLIC PROPERTIES

'Returns 0 if there are no nodes at all. Note that indices for sequential access
'methods are 0-based (thus 0...Count-1).
Public Property Get Count() As Long
    Count = lBinTreeNext
End Property
'_______________________________________________________________________________



'PUBLIC METHODS

'"Add" requests "Bytes" of memory for free usage and associates this memory
'with "Key". Returns the *ADDRESS* of the allocated memory (ideal for CopyMemory).
'Returns 0 if the Key already exists (treat that as an error or read the
'original node ("Item" method) and react accordingly, but never assign data to
'the memory address 0 because your App would undoubtedly crash).
Public Function Add(ByRef Key As String, Bytes As Long) As Long
    Dim lActNode As Long
    Dim lChildNode As Long
    Dim lActChar As Long
    Dim lNewNode As Long
    Dim lStrComp As Long
    Dim bKey() As Byte
    '___________________________________________________________________________
    
    'Copy the string into a Byte array
    bKey = Key
    '___________________________________________________________________________
    
    'If the tree is still empty, create the first root node directly.
    If lBinTreeNext = 0& Then
        'Create the very first node, being the first root.
        lNewNode = MakeNode(sstNoParent, bKey, Bytes)
        'Return it's address to the caller
        Add = vBinTree(lNewNode).Address
        Exit Function
    End If
    '___________________________________________________________________________
    
    'Beginning with the tree's root.
    lActNode = lBinTreeRoot
    
    'Initialize the path vector to trace the nodes we will see. Used to update
    'and balance the binary tree. vPath() was initially dimensioned with 0...9
    'when the object was constructed; later it holds it's highest height.
    'Whenever needed, vPath() is increased by 10 elements (due to tree balancing
    'this won't be the case often: even with 1 million elements we usually get
    'away with less than 40 elements). lPathAct points to the actual element.
    lPathAct = 0&
    
    'Looping until we created a new leaf.
    '(No recursion here please, a simple loop will do the job and contributes
    'much to performance.)
    Do
        'Updating the path we took
        vPath(lPathAct).Index = lActNode
        
        'Comparing new key with existing key (StrComp is significantly faster
        'than two string comparisons with "<" and "=". (Adding 1 million nodes
        'in my test with random 8 char long alpha keys results in 20204798
        'string comparisons.)
        '   < and =:    1000000 Adds: 65.2293464672184 sec net excl overhead
        '   lStrComp:   1000000 Adds: 61.7483442728189 sec net excl overhead
        '(Still this comparison is a bottleneck in mass Add: It takes approx.
        '50% of the time of the whole procedure.)
        '"StrComp" seems to return a Variant ...

        '(StrComp() is slower than CLong(StrComp()) and this slower than my function.)
        'lStrComp = StrComp(Key, vBinTree(lActNode).Key, vbBinaryCompare)
        'lStrComp = CLng(StrComp(Key, vBinTree(lActNode).Key, vbBinaryCompare))
        lStrComp = IStrComp(bKey, vBinTree(lActNode).ByteKey)
        
        'Results in lStrComp: -1 when Key smaller, +1 when larger, 0 when equal.
        'We test ssmdRight first, because *all* keys qualify only here if they
        'are fed in sorted order; otherwise (random feed) they are almost
        'equally distributed. (First branch in Ifs is executed slightly faster.)
        If lStrComp = sstKeyIsLarger Then
            'Right child
            'Updating the path we took
            vPath(lPathAct).Direction = ssmdRight
            
            lChildNode = vBinTree(lActNode).RightTree
            If lChildNode = sstNoChild Then
                'Make a new left child node and link it as a child of the
                'actual node (being the parent now for the new node).
                lNewNode = MakeNode(lActNode, bKey, Bytes)
                vBinTree(lActNode).RightTree = lNewNode
                
                'Heights of the path need to be updated
                UpdateHeights
                
                'Return address to caller and leave
                Add = vBinTree(lNewNode).Address
                Exit Function
            End If
        ElseIf lStrComp = sstKeyIsSmaller Then
            'Left child
            'Updating the path we took
            vPath(lPathAct).Direction = ssmdLeft
            
            lChildNode = vBinTree(lActNode).LeftTree
            If lChildNode = sstNoChild Then
                'Make a new left child node and link it as a child of the
                'actual node (being the parent now for the new node).
                lNewNode = MakeNode(lActNode, bKey, Bytes)
                vBinTree(lActNode).LeftTree = lNewNode
                
                'Heights of the path need to be updated
                UpdateHeights
                
                'Return address to caller and leave
                Add = vBinTree(lNewNode).Address
                Exit Function
            End If
        Else
            'This key already exists.
            'We do not interrupt the client's normal program flow with urging him
            'to react on an error here, but we return 0 and expect him to check
            'for this (invalid) value, except when he is sure that no identical
            'keys do exist.
            Add = 0
            Exit Function
        End If
    
        'Redimensioning the vPath() vector in steps of 10 if needed. Will be
        'executed rarely (in general less than 40 elements for 1 million nodes).
        If lPathAct = lPathMax Then
            lPathMax = lPathMax + NEWELEMENTS
            ReDim Preserve vPath(0& To lPathMax) As rPath
        End If
        lPathAct = lPathAct + 1&
        
        'Following the child's left or right child now.
        lActNode = lChildNode
    Loop
    '___________________________________________________________________________
    
    'Not a valid exit point: We never should arrive here.
    Err.Raise ErrFatalAdd, ErrClass, ErrDesc(ErrFatalAdd)
End Function

'This returns the *index* of the memory associated with the provided key (or -1
'if such a key does not exist). With the index, use the very fast index
'functions to retrieve the associated data (methods "Key", "Address", "Size").
Public Function Item(ByRef Key As String) As Long
    Dim lStrComp As Long
    Dim bKey() As Byte
    '___________________________________________________________________________
    
    'Copy string into byte array
    bKey = Key
    
    'Register this access.
    bLastAcc = bKey
    
    'Begin with root node.
    lLastAcc = lBinTreeRoot
    Do
        'StrComp is significantly faster than two comparisons with "<" and "=".
        '   < and =:    Retrieved all 1000000 nodes by key: 22.78 sec (43896/sec)
        '   StrComp:    Retrieved all 1000000 nodes by key: 18.67 sec (53558/sec)
        'Returns -1 if first string is smaller, 1 if larger, 0 if equal.
        '(Still this comparison is a bottleneck: The following line takes up
        'to 60% of the whole procedure's time. StrComp returns a Variant...)

        '(StrComp() is slower than CLong(StrComp()) and this slower than my function.)
        'lStrComp = StrComp(Key, vBinTree(lLastAcc).Key, vbBinaryCompare)
        'lStrComp = CLng(StrComp(Key, vBinTree(lLastAcc).Key, vbBinaryCompare))
        lStrComp = IStrComp(bKey, vBinTree(lLastAcc).ByteKey)

        If lStrComp = sstKeyIsSmaller Then
            'Continue with left subtree
            lLastAcc = vBinTree(lLastAcc).LeftTree
        ElseIf lStrComp = sstKeyIsLarger Then
            'Continue with right subtree
            lLastAcc = vBinTree(lLastAcc).RightTree
        Else
            '0 as the last case: Lowest chances to hit.
            'Found
            Item = lLastAcc
            'lLastAccPred
            Exit Function
        End If
        If lLastAcc = sstNoChild Then
            'Not existing
            bLastAcc = ""
            Item = sstNoChild
            Exit Function
        End If
    Loop
    'Never should arrive here
    Err.Raise ErrFatalItem, ErrClass, ErrDesc(ErrFatalItem)
End Function

'Accessing the nodes as per sorted key order (lexicographically).
'These methods return the INDICES and *not* the addresses already. thisis because
'the user most likely also wants to access the key, which is cheap when having the
'index. Do *not* store the index: Each additional "Add" might change it drastically.
Public Function Lowest() As Long
    'Begin with root node.
    lLastAcc = lBinTreeRoot
    
    'Follow all left subtrees until there are no more.
    Do While vBinTree(lLastAcc).LeftTree <> sstNoChild
        lLastAcc = vBinTree(lLastAcc).LeftTree
    Loop
    
    'Store as the actual node.
    bLastAcc = vBinTree(lLastAcc).ByteKey
    
    'Return result (INDEX, not address).
    Lowest = lLastAcc
End Function

Public Function Higher() As Long        'Use after Lowest
    'Having an index in lLastAcc we find the successor of that node.
    
    'If lLastAcc has a right child, then it's successor is the minimum in the
    'right subtree.
    If vBinTree(lLastAcc).RightTree <> sstNoChild Then
        'Access right subtree.
        lLastAcc = vBinTree(lLastAcc).RightTree
        'From there follow all left subtrees until there are no more.
        Do While vBinTree(lLastAcc).LeftTree <> sstNoChild
            lLastAcc = vBinTree(lLastAcc).LeftTree
        Loop
        'Store as the actual node.
        bLastAcc = vBinTree(lLastAcc).ByteKey
    Else
        'No right subtree. Find the lowest ancestor of lLastAcc, whose
        'left child is also an ancestor of lLastAcc.
        
        'Follow parent pointers from lLastAcc, until reaching a key larger
        'than lLastAcc's key (if there is none, then lLastAcc is the maximum).
        Do
            lLastAcc = vBinTree(lLastAcc).Parent
            If lLastAcc = sstNoParent Then
                'Is already highest item: No successor is available.
                bLastAcc = ""
                Exit Do
            ElseIf IStrComp(vBinTree(lLastAcc).ByteKey, bLastAcc) = sstKeyIsLarger Then
                'Found: Store as the actual node (just because it becomes cheap).
                bLastAcc = vBinTree(lLastAcc).ByteKey
                Exit Do
            End If
        Loop
        'lLastAcc is either the successor or -1 (already highest item)
    End If

    'Return result (INDEX, not address).
    Higher = lLastAcc
End Function

Public Function Highest() As Long
    'Begin with root node.
    lLastAcc = lBinTreeRoot
    
    'Follow all right subtrees until there are no more.
    Do While vBinTree(lLastAcc).RightTree <> sstNoChild
        lLastAcc = vBinTree(lLastAcc).RightTree
    Loop
    
    'Store as the actual node.
    bLastAcc = vBinTree(lLastAcc).ByteKey
    
    'Return result (INDEX, not address).
    Highest = lLastAcc
End Function

Public Function Lower() As Long         'Use after Highest
    'Having an index in lLastAcc we find the predecessor of that node.
    
    'If lLastAcc has a left child, then it's predecessor is the maximum in the
    'left subtree.
    If vBinTree(lLastAcc).LeftTree <> sstNoChild Then
        'Access left subtree.
        lLastAcc = vBinTree(lLastAcc).LeftTree
        'From there follow all right subtrees until there are no more.
        Do While vBinTree(lLastAcc).RightTree <> sstNoChild
            lLastAcc = vBinTree(lLastAcc).RightTree
        Loop
        'Store as the actual node (just because it becomes cheap).
        bLastAcc = vBinTree(lLastAcc).ByteKey
    Else
        'No left subtree. Find the lowest ancestor of lLastAcc, whose
        'right child is also an ancestor of lLastAcc.
        
        'Follow parent pointers from lLastAcc, until reaching a key smaller
        'than lLastAcc's key (if there is none, then lLastAcc is the minimum).
        Do
            lLastAcc = vBinTree(lLastAcc).Parent
            If lLastAcc = sstNoParent Then
                'Is already smallest item: No predecessor is available.
                bLastAcc = ""
                Exit Do
            ElseIf IStrComp(vBinTree(lLastAcc).ByteKey, bLastAcc) = sstKeyIsSmaller Then
                'Found: Store as the actual node (just because it becomes cheap).
                bLastAcc = vBinTree(lLastAcc).ByteKey
                Exit Do
            End If
        Loop
        'lLastAcc is either the successor or -1 (already highest item)
    End If

    'Return result (INDEX, not address).
    Lower = lLastAcc
End Function



'Accessing the nodes by their insertion order (0=first, lBinTreeNext-1=last).
'Returns the address of the node, or 0 if there is no such index.
Public Function Address(Index As Long) As Long
    If Index >= lBinTreeNext Then
        Address = 0
    ElseIf Index < 0 Then
        Address = 0
    Else
        Address = vBinTree(Index).Address
    End If
End Function

'Returns the reserved memory of the node 8in bytes), or 0 if there is no such index.
Public Function Size(Index As Long) As Long
    If Index >= lBinTreeNext Then
        Size = 0
    ElseIf Index < 0 Then
        Size = 0
    Else
        Size = vBinTree(Index).Bytes
    End If
End Function

'Returns the key of the node, or a null-length string ("") if there is no
'such index.
Public Function Key(Index As Long) As String
    If Index >= lBinTreeNext Then
        Key = ""
    ElseIf Index < 0 Then
        Key = ""
    Else
        'Return as a string, user is not interested in a Byte Array
        Key = vBinTree(Index).ByteKey
    End If
End Function
'_______________________________________________________________________________



'INTERNAL ROUTINES

'Creating a new node below a given parent. Returns the index into vBinTree().
Private Function MakeNode(ParentIndex As Long, ByRef Key() As Byte, Bytes As Long) As Long
    'If this new node requires that we make new elements to hold more nodes,
    'we expand our vector by doubling it (a constant number of elements *always*
    'becomes a bottleneck at some point of time, even if it's 1000).
    If lBinTreeNext > lBinTreeMax Then
        lBinTreeMax = lBinTreeMax * 2
        ReDim Preserve vBinTree(0 To lBinTreeMax) As rNode
    End If
    
    'lBinTreeNext is the element index into the vector.
    With vBinTree(lBinTreeNext)
        .ByteKey = Key
        .Address = ifMalloc.Alloc(Bytes)
        .Bytes = Bytes
        .Parent = ParentIndex
        .LeftTree = sstNoChild
        .RightTree = sstNoChild
    End With
    
    'Return the index as this function's result
    MakeNode = lBinTreeNext
    
    'Prepare for next node.
    lBinTreeNext = lBinTreeNext + 1&
End Function



'BALANCING THE TREE
'Depending on the nodes' actual balances, one of four balancing activities can
'be called (BalanceLL, BalanceRR, BalanceLR or BalanceRL). In a test scenario
'with 1 million nodes with 8 char long RANDOM keys consisting of all upper- and
'lowercase alphabetic characters A...Z and a...z, the balancing routines were
'called so many times:          With 1 mill already SORTED input (key length 4):
'    Rebalances LL 119628            0
'    Rebalances RR 118737       999980
'    Rebalances LR 114924            0
'    Rebalances RL 115005            0
'    Total         468294       999980
Private Sub UpdateHeights()
    Dim lLevel As Long
    Dim lNode As Long
    Dim lLeftTree As Long
    Dim lRightTree As Long
    Dim iUpperBalance As Integer        'Integer to match UDT definition
    Dim iLowerBalance As Integer        'Integer to match UDT definition
    Dim iNewHeight As Integer           'Integer to match UDT definition
    '___________________________________________________________________________
    
    'Must be reverse height, because we might break update for upper levels.
    For lLevel = lPathAct To 0& Step -1&
        lNode = vPath(lLevel).Index
        
        'Increase height along path. Direction is either ssmdLeft or ssmdRight.
        'If the Keys are fed sorted, only ssmdRight will appear (if fed in
        'random order, there are almost equal chances for the left and the right
        'path): For this reason, ssmdRight is tested first (slightly faster).
        If vPath(lLevel).Direction = ssmdRight Then
            iNewHeight = vBinTree(lNode).RightHeight + 1&
            vBinTree(lNode).RightHeight = iNewHeight
            'If right was *not* responsible for a height increase, this breaks
            'the update (no propagation farther up).
            If iNewHeight <= vBinTree(lNode).LeftHeight Then
                Exit For
            End If
        Else
            iNewHeight = vBinTree(lNode).LeftHeight + 1&
            vBinTree(lNode).LeftHeight = iNewHeight
            'If Left was *not* responsible for a height increase, this breaks
            'the update (no propagation farther up).
            If iNewHeight <= vBinTree(lNode).RightHeight Then
                Exit For
            End If
        End If
    
        'Nodes with a too heavy weight left or right cause a tree balancing
        'activity now. Too heavy is, when there is a height difference of +/-2.
        iUpperBalance = vBinTree(lNode).RightHeight - vBinTree(lNode).LeftHeight
        
        'testing right-heavy first because we want RR to be executed as fast as
        'possible to acknowledge that RR type is executed almost always when keys
        'are fed already sorted to Add; whereas the distribution of the 4 types
        'is almost the same (totalling in about 44%) when the keys are fed in
        'random order.
        If iUpperBalance = ssmbRightHeavy Then
            'Node is right-heavy (+2). Examine it's right child.
            lRightTree = vBinTree(lNode).RightTree
            iLowerBalance = vBinTree(lRightTree).RightHeight - vBinTree(lRightTree).LeftHeight

            'Testing ssmbRightBalanced first to take care for the case that
            'keys may come in already sorted, in which case 99.8% of all Adds
            'require a RR balancing. If keys are fed random, all 4 types are
            'equally distributed.
            If iLowerBalance = ssmbRightBalanced Then
                'Left child is right-balanced. RR situation.
                BalanceRR lNode
            Else
                'Cannot be ssmbBalanced, and so must be ssmbLeftBalanced.
                BalanceRL lNode
            End If
            'This breaks updates further back!
            Exit For
        ElseIf iUpperBalance = ssmbLeftHeavy Then
            'Node is left-heavy (-2). Examine it's left child.
            lLeftTree = vBinTree(lNode).LeftTree
            iLowerBalance = vBinTree(lLeftTree).RightHeight - vBinTree(lLeftTree).LeftHeight

            If iLowerBalance = ssmbRightBalanced Then
                'Left child is right-balanced. LR situation.
                BalanceLR lNode
            Else
                'Cannot be ssmbBalanced, and so must be ssmbLeftBalanced.
                BalanceLL lNode
            End If
            'This breaks updates further back!
            Exit For
        End If
    Next lLevel
End Sub


'_______________________________________________________________________________
'
'INTERNAL ROUTINES TO BALANCE THE BINARY TREE

'LL - A is the node that the rotation is performed on. This rotation is performed
'when A is unbalanced to the left (the left subtree is 2 higher than the right
'subtree) and B is leftheavy (the left subtree of B is 1 higher than the right
'subtree of B). T1, T2 and T3 represent subtrees (a node was added to T1 which
'made B leftheavy and unbalanced A). P is A's parent and listed only because A's
'relationship with P is taken over by B, making relinks neccessary.
'
'        P                        P
'        |                        |
'        A                        B
'       / \      ---------->     / \
'      B  T3                    T1  A
'     / \                       ** / \
'    T1 T2                        T2 T3
'    **
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |B |  |P |T2|  |B |
'|B |  |T2|A |  |A |P |
'|T2|  |  |B |  |  |A |
'|P |  A  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceLL(lA As Long)
    Dim lB As Long
    Dim lP As Long
    Dim lT2 As Long
    
    'Indices of old state
    lB = vBinTree(lA).LeftTree
    lT2 = vBinTree(lB).RightTree
    lP = vBinTree(lA).Parent
    
    'New links
    vBinTree(lA).LeftTree = lT2
    If lT2 <> sstNoChild Then vBinTree(lT2).Parent = lA
    vBinTree(lB).RightTree = lA
    vBinTree(lA).Parent = lB
    If lP <> sstNoParent Then
        'A was either P's left or right child, B is taking that place now.
        vBinTree(lB).Parent = lP
        If vBinTree(lP).LeftTree = lA Then
            vBinTree(lP).LeftTree = lB
        Else
            vBinTree(lP).RightTree = lB
        End If
    Else
        'If A was the whole tree's root, then it's role is taken over by B now.
        vBinTree(lB).Parent = sstNoParent
        lBinTreeRoot = lB
    End If
    
    'New weights
    vBinTree(lA).LeftHeight = vBinTree(lA).LeftHeight - 2
    vBinTree(lB).RightHeight = vBinTree(lB).RightHeight + 1
End Sub
'
'RR - A is the node that the rotation is performed on. This rotation is performed
'when A is unbalanced to the right (the right subtree is 2 higher than the left
'subtree) and B is rightheavy (the right subtree of B is 1 higher than the left
'subtree of B). T1, T2 and T3 represent subtrees (a node was added to T3 which
'made B rightheavy and unbalanced A). P is A's parent and listed only because A's
'relationship with P is taken over by B, making relinks neccessary.
'(This rotation is performed on almost every node if the keys are added in already
'sorted order, and no other rotations are performed at all then.)
'
'        P                        P
'        |                        |
'        A                        B
'       / \      ---------->     / \
'      T1  B                    A  T3
'         / \                  / \ **
'        T2  T3               T1 T2
'            **
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |  |B |P |  |T2|B |
'|B |T2|  |A |A |  |P |
'|T2|  |  |B |  |  |A |
'|P |  A  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceRR(lA As Long)
    Dim lB As Long
    Dim lP As Long
    Dim lT2 As Long
    
    'Indices of old state
    lB = vBinTree(lA).RightTree
    lP = vBinTree(lA).Parent
    lT2 = vBinTree(lB).LeftTree
    
    'New links
    vBinTree(lA).RightTree = lT2
    If lT2 <> sstNoChild Then vBinTree(lT2).Parent = lA
    vBinTree(lB).LeftTree = lA
    vBinTree(lA).Parent = lB
    If lP <> sstNoParent Then
        'A was either P's left or right child, B is taking that place now.
        vBinTree(lB).Parent = lP
        If vBinTree(lP).LeftTree = lA Then
            vBinTree(lP).LeftTree = lB
        Else
            vBinTree(lP).RightTree = lB
        End If
    Else
        'If A was the whole tree's root, then it's role is taken over by B now.
        vBinTree(lB).Parent = sstNoParent
        lBinTreeRoot = lB
    End If
    
    'New weights
    vBinTree(lA).RightHeight = vBinTree(lA).RightHeight - 2
    vBinTree(lB).LeftHeight = vBinTree(lB).LeftHeight + 1
End Sub
'
'LR - C is the node that the rotation is performed on. This rotation is performed
'when C is unbalanced to the left (the left subtree is 2 higher than the right
'subtree), A is rightheavy (the right subtree of A is 1 higher than the left
'subtree of A) and B is leftheavy. T1, T2, T3, and T4 represent subtrees (a node
'was added to T2 which made B leftheavy, made A rightheavy and unbalanced C).
'This consists of a single left rotation at node A, followed by a single right
'at node C. P is C's parent and listed only because C's relationship with P is
'taken over by B, making relinks neccessary.
'
'        P                         P
'        |                         |
'        C                         B
'       / \      ---------->     /   \
'      A  T4                    A     C
'     / \                      / \   / \
'    T1  B                    T1 T2 T3 T4
'       / \                      **
'      T2 T3
'      **
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |  |B |C |  |T2|B |
'|B |T2|T3|A |A |C |P |
'|C |A |  |P |T3|  |B |
'|T2|  |  |B |  |  |A |
'|T3|  |  |B |  |  |C |
'|P |  C  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceLR(lC As Long)
    Dim lA As Long
    Dim lB As Long
    Dim lT2 As Long
    Dim lT3 As Long
    Dim lP As Long
    
    'Indices of old state
    lA = vBinTree(lC).LeftTree
    lB = vBinTree(lA).RightTree
    lP = vBinTree(lC).Parent
    lT2 = vBinTree(lB).LeftTree
    lT3 = vBinTree(lB).RightTree

    'New links
    vBinTree(lA).RightTree = lT2
    vBinTree(lA).Parent = lB
    vBinTree(lB).LeftTree = lA
    vBinTree(lB).RightTree = lC
    vBinTree(lC).LeftTree = lT3
    vBinTree(lC).Parent = lB
    If lT2 <> sstNoChild Then vBinTree(lT2).Parent = lA
    If lT3 <> sstNoChild Then vBinTree(lT3).Parent = lC
    If lP <> sstNoParent Then
        'C was either P's left or right child, B is taking that place now.
        vBinTree(lB).Parent = lP
        If vBinTree(lP).LeftTree = lC Then
            vBinTree(lP).LeftTree = lB
        Else
            vBinTree(lP).RightTree = lB
        End If
    Else
        'If C was the whole tree's root, then it's role is taken over by B now.
        vBinTree(lB).Parent = sstNoParent
        lBinTreeRoot = lB
    End If
    
    'New weights
    vBinTree(lA).RightHeight = vBinTree(lA).RightHeight - 1
    vBinTree(lB).LeftHeight = vBinTree(lB).LeftHeight + 1
    vBinTree(lB).RightHeight = vBinTree(lB).RightHeight + 1
    vBinTree(lC).LeftHeight = vBinTree(lC).LeftHeight - 2
End Sub
'
'RL - A is the node that the rotation is performed on. This rotation is performed
'when A is unbalanced to the right (the right subtree is 2 higher than the left
'subtree), C is leftheavy (the left subtree of A is 1 higher than the right
'subtree of A) and B is rightheavy. T1, T2, T3, and T4 represent subtrees (a node
'was added to T3 which made B rightheavy, made C leftheavy and unbalanced A).
'This consists of a single right at node C, followed by a single left at node A.
'P is A's parent and listed only because A's relationship with P is taken over
'by B, making relinks neccessary.
'
'        P                         P
'        |                         |
'        A                         B
'       / \      ---------->     /   \
'      T1  C                    A     C
'         / \                  / \   / \
'        B  T4                T1 T2 T3 T4
'       / \                         **
'      T2 T3
'         **
'+--+--------+--------+
'|  | Before | After  |
'+--+--+--+--+--+--+--+
'|  |LT|RT|PA|LT|RT|PA|
'+--+--+--+--+--+--+--+
'|A |  |C |P |  |T2|B |
'|B |T2|T3|C |A |C |P |
'|C |B |  |A |T3|  |B |
'|T2|  |  |B |  |  |A |
'|T3|  |  |B |  |  |C |
'|P |  A  |  |  B  |  |
'+--+--+--+--+--+--+--+
'
Private Sub BalanceRL(lA As Long)
    Dim lB As Long
    Dim lC As Long
    Dim lT2 As Long
    Dim lT3 As Long
    Dim lP As Long
    
    'Indices of old state
    lP = vBinTree(lA).Parent
    lC = vBinTree(lA).RightTree
    lB = vBinTree(lC).LeftTree
    lT2 = vBinTree(lB).LeftTree
    lT3 = vBinTree(lB).RightTree
    
    'New links
    vBinTree(lA).RightTree = lT2
    vBinTree(lA).Parent = lB
    vBinTree(lB).LeftTree = lA
    vBinTree(lB).RightTree = lC
    vBinTree(lC).LeftTree = lT3
    vBinTree(lC).Parent = lB
    If lT2 <> sstNoChild Then vBinTree(lT2).Parent = lA
    If lT3 <> sstNoChild Then vBinTree(lT3).Parent = lC
    If lP <> sstNoParent Then
        'A was either P's left or right child, B is taking that place now.
        vBinTree(lB).Parent = lP
        If vBinTree(lP).LeftTree = lA Then
            vBinTree(lP).LeftTree = lB
        Else
            vBinTree(lP).RightTree = lB
        End If
    Else
        'If A was the whole tree's root, then it's role is taken over by B now.
        vBinTree(lB).Parent = sstNoParent
        lBinTreeRoot = lB
    End If
    
    'New weights
    vBinTree(lA).RightHeight = vBinTree(lA).RightHeight - 2
    vBinTree(lB).LeftHeight = vBinTree(lB).LeftHeight + 1
    vBinTree(lB).RightHeight = vBinTree(lB).RightHeight + 1
    vBinTree(lC).LeftHeight = vBinTree(lC).LeftHeight - 1
End Sub
'_______________________________________________________________________________

'Searching for a fast replacement of the slow StrComp.
Private Function IStrComp(ByRef bStr() As Byte, ByRef bCmp() As Byte) As Long
'    Dim lSameChars As Long
'
'(As per MSDN October 2002, RTLCompareMemory() is only included in Win2K and XP.)
'Private Declare Function CompareMemory Lib "ntdll" Alias "RtlCompareMemory" _
'    (buffA As Any, buffB As Any, ByVal Length As Long) As Long

'    'CompareMemory returns how many byte elements are the same.
'    lSameChars = CompareMemory(bStr(0), bCmp(0), UBound(bStr))
'    If bStr(lSameChars) > bCmp(lSameChars) Then
'        IStrComp = sstKeyIsLarger
'        Exit Function
'    ElseIf bStr(lSameChars) < bCmp(lSameChars) Then
'        IStrComp = sstKeyIsSmaller
'        Exit Function
'    End If
'    'Else returning 0 (doing nothing as it's the default result)
    
    Dim lPos As Long

    For lPos = 0 To UBound(bStr) Step 2&     'Step 2 for ANSI test, else 1
        'bCmp might end sooner than bStr, but be otherwise equal. In this
        'case bStr is larger.
        If lPos > UBound(bCmp) Then
            IStrComp = sstKeyIsSmaller
            Exit Function
        'Instead of comparing each char twice (once for larger and once for
        'smaller), we compare once for inequality. Only if not equal we do a
        'final test on how it was not equal. ("<>" is faster than ">" or "<"
        'alone, so even more when testing both.)
        ElseIf bStr(lPos) <> bCmp(lPos) Then
            If bStr(lPos) > bCmp(lPos) Then
                IStrComp = sstKeyIsLarger
            Else
                IStrComp = sstKeyIsSmaller
            End If
            Exit Function
        End If
    Next lPos
    'Strings are equal if Cmp ends as well, or Cmp is longer and thus larger.
    If UBound(bStr) <> UBound(bCmp) Then IStrComp = sstKeyIsLarger
    'Else implicitely returning sstKeyIsEqual = 0
End Function
