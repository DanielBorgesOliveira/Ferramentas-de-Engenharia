; Bind (insert) all xrefs in the drawing.
(defun c:BindXrefs ( / xrefs xrefName xrefPath)
	;; Ask the user once at the beginning with predefined keywords
	(initget "Sim Não") ;; Initialize the keywords
	(setq userResponse 
		(getkword "\nGostaria de converter todas referências externas (XREF) para bloco? [Sim/Não] <Sim>: "))

	;; Default response is "Sim"
	(if (or (= (strcase userResponse) "SIM") (not userResponse))
		(progn
			;; Get the list of all Xrefs in the current drawing
			(setq xrefs (vla-get-blocks (vla-get-activedocument (vlax-get-acad-object))))

			;; Iterate through the Xrefs and bind them
			(vlax-for xref xrefs
				(if (= :vlax-true (vla-get-isxref xref))
					(progn
						;; Get the name of the Xref
						(setq xrefName (vla-get-name xref))
						
						;; Temporarily detach the Xref path annotation (if any)
						(setq xrefPath (getenv "ACADXREFDISPLAY")) ;; Backup original value
						(setenv "ACADXREFDISPLAY" "0") ;; Disable Xref path display
						
						;; Bind the Xref
						(command "_.-XREF" "_B" xrefName "")
						
						;; Restore Xref path display setting
						(setenv "ACADXREFDISPLAY" xrefPath) ;; Restore original value
						
						(princ (strcat "\nXref " xrefName " bound successfully."))
					)
				)
			)
			(princ "\nTodas as referências externas foram convertidas em blocos.")
		)
		;; If the user selects "Não", exit without binding
		(progn
			(princ "\nNenhuma referência externa foi convertida. Processo cancelado pelo usuário.")
		)
	)
	(princ)
)