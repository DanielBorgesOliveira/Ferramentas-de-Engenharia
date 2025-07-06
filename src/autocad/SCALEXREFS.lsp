(defun c:SCALEXREFS ()
  (setq ss (ssget '((0 . "INSERT")))) ; Select all Xrefs (blocks with INSERT)
  (if ss
    (progn
      (setq ref_xref (entsel "\nSelect reference Xref: ")) ; Ask for reference Xref
      (if ref_xref
        (progn
          (setq ref_ent (car ref_xref)
                ref_data (entget ref_ent)
                ref_scale_x (cdr (assoc 41 ref_data)) ; X scale
                ref_scale_y (cdr (assoc 42 ref_data)) ; Y scale
                ref_scale_z (cdr (assoc 43 ref_data))) ; Z scale
          
          (setq i 0)
          (while (< i (sslength ss))
            (setq ent (ssname ss i))
            (setq ent_data (entget ent))
            (setq ent_scale_x (cdr (assoc 41 ent_data))) ; Get current X scale
            
            (if (/= ent_scale_x ref_scale_x) ; If different, update all scales
              (progn
                (setq new_data (subst (cons 41 ref_scale_x) (assoc 41 ent_data) ent_data))
                (setq new_data (subst (cons 42 ref_scale_y) (assoc 42 new_data) new_data))
                (setq new_data (subst (cons 43 ref_scale_z) (assoc 43 new_data) new_data))
                (entmod new_data) ; Apply changes
                (princ (strcat "\nXref scaled to match reference: " (cdr (assoc 2 ent_data))))
              )
            )
            (setq i (1+ i))
          )
        )
      )
    )
  )
  (princ "\nNo Xrefs found.")
  (princ)
)
