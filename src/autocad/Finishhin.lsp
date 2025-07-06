; Initializes AutoCAD's Visual LISP COM interface.
(vl-load-com)

; Used to suppress unnecessary output or to clean up the command line after executing a script.
(princ)

(defun c:RenameLayoutToDocName ( / docname layoutname)
	;; Rnames the current layout to document name.
  (vl-load-com)

  ;; Get the document name without path or extension
  (setq docname (vl-filename-base (getvar 'DWGNAME)))

  ;; Get current layout name
  (setq layoutname (getvar 'CTAB))

  ;; Do not rename Model tab
  (if (equal layoutname "Model")
    (progn
      (princ "\nCannot rename the 'Model' layout.")
    )
    (progn
      ;; Rename the layout
      (vla-put-name
        (vla-item
          (vla-get-layouts
            (vla-get-activedocument (vlax-get-acad-object))
          )
          layoutname
        )
        docname
      )
      (princ (strcat "\nLayout renamed to: " docname))
    )
  )

  (princ)
)


(defun zea ( / acapp acdoc aclay )
	; Zoom extents in all layouts.
    (setq acapp (vlax-get-acad-object)
          acdoc (vla-get-activedocument acapp)
          aclay (vla-get-activelayout acdoc)
    )
    (vlax-for layout (vla-get-layouts acdoc)
        (vla-put-activelayout acdoc layout)
        (if (eq acpaperspace (vla-get-activespace acdoc))
            (vla-put-mspace acdoc :vlax-false)
        )
        (vla-zoomextents acapp)
    )
    (vla-put-activelayout acdoc aclay)
    (princ)
)

(defun c:BindXrefs ( / xrefs xrefName xrefPath)
	; Bind (insert) all xrefs in the drawing.
	; TODO: verificar porque o caminho do xref está aparecendo em forma de texto no bloco.

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
						;;(setenv "ACADXREFDISPLAY" xrefPath) ;; Restore original value
						
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

(defun ChangeBlockSpace ( blockName )
	; Change space of a block
	
	; Validate the block name is provided
	(if (not blockName)
		(progn
			(setq blockName (getstring "\nEnter block name: "))
		)
	)
	;(princ (strcat "DEBUG **** " blockName " *****"))
	(setq ss (ssget "X" (list '(0 . "INSERT") (cons 2 blockName))))
	(command "_CHSPACE" ss "")
)

(defun DeleteLayerContent (layer / ss i ent blks obj)
	; Delete comment layers and its contents.
    (vl-load-com)
    (if (tblsearch "LAYER" layer) ; Check if the layer exists
        (progn
            (setq ss (ssget "_X" (list (cons 8 layer)))) ; Select all objects on the layer
            (if ss ; Proceed only if the selection set is not nil
                (progn
                    (setq i -1)
                    (while (setq ent (ssname ss (setq i (1+ i))))
                        (entdel ent)
					) ; Delete each object
                )
            )
            (vlax-for blks (vla-get-Blocks
				(vla-get-ActiveDocument
				(vlax-get-acad-object)))
                ; Delete blocks on the layer
				(vlax-for obj blks
                    (if (eq (strcase layer) (strcase (vla-get-layer obj)))
						(vla-delete obj)
					)
				)
            )
        )
        (prompt "\nLayer does not exist.")
    )
	(princ)
)

(defun c:Line2PlineJoin (/ ss i en elist)
	; Convert all lines to pline.
	
  ;; Initialize the iterator
  (setq i 0)
  ;; Get selection set of all lines in the drawing
  (setq ss (ssget "X" '((0 . "LINE"))))
  (if ss
    (progn
      ;; Join all polylines
      (command "_PEDIT" "_M" ss "" "_Y" "_JOIN" 0.02 "")
      (princ "\nConverted lines to polyline and joined them.")
    )
    (princ "\nNo lines found in the drawing.")
  )
  (princ)
)

(defun c:TextToMText (/ ss i en ed)
	; Convert all text to mtext.
	
  ;; Create a selection set of all text objects
  (setq ss (ssget "X" '((0 . "TEXT"))))
  (if ss
    (progn
      ;; Loop through all text entities in the selection set
      (setq i 0)
      (repeat (sslength ss)
        (setq en (ssname ss i)) ; Get entity name
        (setq ed (entget en))   ; Get entity data
        ;; Create MTEXT entity data
        (entmake
          (list
            '(0 . "MTEXT")
            (assoc 10 ed) ; Insertion point
            (assoc 40 ed) ; Text height
            (assoc 1 ed)  ; Text value
            (assoc 41 ed) ; Width factor (if exists)
            (assoc 7 ed)  ; Text style (if exists)
            '(71 . 1)     ; Attachment point (1 = Top left)
            '(72 . 1)     ; Drawing direction (1 = Left to right)
          )
        )
        ;; Delete the original TEXT entity
        (entdel en)
        (setq i (1+ i))
      )
    )
    (alert "No text entities found in the drawing.")
  )
  (princ)
)

(defun c:ExplodeAllTables ()
	; Explode all table in the drawing.
  (vl-load-com)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object))) ; Get the active document
  (setq modelSpace (vla-get-ModelSpace doc)) ; Get the Model Space
  
  (vlax-for obj modelSpace
    (if (eq (vla-get-ObjectName obj) "AcDbTable") ; Check if the object is a table
      (progn
        (command "_explode" (vlax-vla-object->ename obj)) ; Explode the table
        (princ "Table exploded successfully.\n")
      )
    )
  )

  (princ "All tables have been exploded.\n") ; Confirm completion
  (princ)
)

(defun c:FreezeLayer ( layName /  layTblRec)
  (if (tblsearch "LAYER" layName)
    (progn
      (setq layTblRec (tblobjname "LAYER" layName))
      (if layTblRec
        (progn
          (setq layTblRec (entget layTblRec))
          (if (/= (cdr (assoc 70 layTblRec)) 1) ; already frozen?
            (progn
              (setq layTblRec (subst (cons 70 1) (assoc 70 layTblRec) layTblRec))
              (entmod layTblRec)
              (princ (strcat "\nLayer '" layName "' has been frozen."))
            )
            (princ "\nLayer is already frozen.")
          )
        )
        (princ "\nCould not retrieve layer object.")
      )
    )
    (princ "\nLayer does not exist.")
  )
  (princ)
)

(defun c:FinishHim (/ ss)
	(zea)
	
	(DeleteLayerContent "1 - Comentarios Abertos")
	(DeleteLayerContent "1 - Comentarios Resolvidos")
	
	; Freeze our layer with tags.
	(c:FreezeLayer "1-TagNosso")
	
	; Bind all xrefs.
	(c:BindXrefs)
	
	; Remove all datalink
	(dictremove (namedobjdict) "ACAD_DATALINK")
	
	; Explode all tables
	;(c:ExplodeAllTables)
	
	; Move to paper space.
	(command "TILEMODE" "0")
	
	; Rename the current layout to  document's name.
	(c:RenameLayoutToDocName)
	
	; Acess model space
	(command "mspace")
	(command "-dwgunits" "3" "2" "4" "N" "N")
	
	; Call units to set the units of the drawing in milimeters.
	(command "_.units" "2" "4" "1" "2" "0.00" "N")
	
	; Call purge to remove unused entities.
	(command "_.purge" "A" "*" "N")
	(command "\_.purge" "A" "\*" "N")
	(command "\_.purge" "A" "\*" "N")
	
	; Call overkill to remove overlapping lines.
	(command "\_.-overkill" "_all" "" "\_done")
	
	; Controls which properties are selected for the SETBYLAYER command. 1 = color.
	(command "setbylayermode" "1")
	
	; Changes the property overrides of selected objects to ByLayer, except blocks.
	(command "setbylayer" "_all" "" "y" "n")
	
	; Audit the document.
	(command "_.audit" "Yes" "")
	
	; Regenerate the  document.
	(command "_.regenall")
	
	(c:Line2PlineJoin)
	
	; Move all blocks to front
	(setq ss (ssget "X" '((0 . "INSERT"))))
	(command "_.draworder" ss "" "_front")
	
	; Move the client format to paper space.
	(ChangeBlockSpace "A1")
	
	; Move to paper space
	(command "pspace")
	
	; Save the documento
	;(command "qsave")
); defun