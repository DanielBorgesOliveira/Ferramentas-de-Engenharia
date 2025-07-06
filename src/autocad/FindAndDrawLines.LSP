(defun c:FindAndDrawLines (/ keywords origin ss i text insPt)
  (setq keywords '(
    "FV"
    "FO"
	"PCV"
	"LE"
	"FE"
  ))
  
  (setq origin '(-541 -541 0)) ; Origin point
  (foreach keyword keywords
    (setq ss (ssget "X" (list (cons 0 "TEXT") (cons 1 (strcat "*" keyword "*")))))
    (if ss
      (progn
        (setq i 0)
        (while (< i (sslength ss))
          (setq text (ssname ss i))
          (setq insPt (cdr (assoc 10 (entget text)))) ; Get insertion point
          (command "LINE" origin insPt "") ; Draw line
          (setq i (1+ i))
        )
      )
    )
  )
  (princ)
)
(c:FindAndDrawLines)



