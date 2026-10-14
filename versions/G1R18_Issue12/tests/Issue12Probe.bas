Attribute VB_Name = "Issue12Probe"
Option Explicit
' Test-copy-only probe. Never saved into a delivery workbook.
Public Function Issue12State() As String
    Issue12State = "last_error=" & RPX_LastError & _
        ";lower=" & RBestLower.value & ";upper=" & RBestUpper.value & _
        ";lower_q=" & RBestLower.q & ";upper_q=" & RBestUpper.q & _
        ";lower_complete=" & RBestLower.complete & ";upper_complete=" & RBestUpper.complete & _
        ";lower_root=" & RBestLower.lowerRootValid & ";upper_root=" & RBestUpper.upperRootValid & _
        ";lower_elements=" & RBestLower.elements & ";upper_elements=" & RBestUpper.elements & _
        ";pair_complete=" & RPX_AdaptPairComplete() & ";search_incomplete=" & RFsSearchIncomplete & _
        ";original_ready=" & XOriginalCompiled.ready & ";original_q=" & XOriginalCompiled.subdivisions & _
        ";original_mode=" & XOriginalCompiled.modeName & ";q_now=" & RSubdivisions & ";input_epoch=" & RPX_InputEpoch
End Function
