; Acess model space
(command "mspace")
(command "TILEMODE" "1") ; Move to model space.
(command "ZOOM" "E") ; Zoom extents
(command "UCS" "_W") ; Reset UCS to original position
(command "-dwgunits" "3" "2" "4" "YES" "YES" "NO") ; Call units to set the units of the drawing in milimeters.
(command "_.units" "2" "4" "1" "2" "0.00" "N")
(command "DATALINKUPDATE" "_U" "_K")
;(command "_ATTSYNC" "Name" "A1")
(command "TILEMODE" "0") ; Move to paper space.
; Rename the current layout to  document's name.
(c:RenameLayoutToDocName)
(command "ZOOM" "E") ; Zoom extents
(command "UCS" "_W") ; Reset UCS to original position
(command "mspace") ; Acess model space
(setvar "CLAYER" "0") ; Set the layer "0" as active.
;(c:rp)
;(c:FinishHim)
;(command "qsave")
;(command "close")