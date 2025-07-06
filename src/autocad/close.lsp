(defun c:CC ()
	(command "TILEMODE" "1") ; Switch to Model Space
	(command "zoom" "e") ; Zoom extents
	
	(command "TILEMODE" "0") ; Switch to Paper Space (Layout1)
	(command "zoom" "e") ; Zoom extents
	
	(command "_.QSAVE" "Y") ; Save the drawing
	(command "_.CLOSE") ; Close the drawing
	(princ)
)