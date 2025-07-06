(defun c:SumDims ( / ss i dimObj dimVal total)
  (setq ss (ssget '((0 . "DIMENSION")))) ; Select all dimension objects
  (if ss
    (progn
      (setq total 0) ; Initialize total
      (setq i 0) ; Initialize index
      (while (< i (sslength ss))
        (setq dimObj (vlax-ename->vla-object (ssname ss i))) ; Get dimension object
        (setq dimVal (vla-get-Measurement dimObj)) ; Get dimension value
        (setq total (+ total dimVal)) ; Add to total
        (setq i (1+ i)) ; Increment index
      )
      (princ (strcat "\nTotal of dimensions: " (rtos total 2 2))) ; Display total
    )
    (princ "\nNo dimensions selected.") ; No dimensions found
  )
  (princ) ; Exit quietly
)
