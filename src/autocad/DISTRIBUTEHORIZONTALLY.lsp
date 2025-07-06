(defun c:DISTRIBUTEHORIZONTALLY (/ ss num spacing x-offset i ent ename minPt maxPt minx maxx width)
  (setq ss (ssget))  ; Select objects
  (if ss
    (progn
      (setq num (sslength ss)
            spacing 500  ; Adjust spacing as needed
            x-offset 0
            i 0)
      
      (while (< i num)
        (setq ename (ssname ss i)
              ent (vlax-ename->vla-object ename))  ;; Convert entity to vla object
        
        ;; Get bounding box safely
        (setq minPt (vlax-make-safearray vlax-vbDouble '(0 . 2)))
        (setq maxPt (vlax-make-safearray vlax-vbDouble '(0 . 2)))
        
        (if (not (vl-catch-all-error-p (setq result (vl-catch-all-apply 'vla-getboundingbox (list ent 'minPt 'maxPt)))))
          (progn
            ;; Convert safe array to lists
            (setq minx (vlax-safearray->list minPt))
            (setq maxx (vlax-safearray->list maxPt))

            (if (and minx maxx)  ;; Ensure valid coordinates were found
              (progn
                (setq width (- (car maxx) (car minx)))  ;; Calculate object width
                ;; Move object to new position
                (command "_.move" ename "" "_non" minx "_non" (list x-offset (cadr minx)))
                ;; Update offset for next object
                (setq x-offset (+ x-offset width spacing))
              )
            )
          )
        )
        (setq i (1+ i))
      )
    )
    (princ "\nNo objects selected.")
  )
  (princ)
)
