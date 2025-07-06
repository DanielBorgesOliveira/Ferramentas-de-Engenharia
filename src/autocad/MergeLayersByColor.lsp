(defun c:MergeLayersByColor ()
  (setq layerTable (tblnext "layer" T))  ; Start with the first layer in the drawing
  (setq colorMap '())  ; Create an empty association list to store layers by color
  
  ; Iterate through all layers in the drawing
  (while layerTable
    (setq layerName (cdr (assoc 2 layerTable)))  ; Get the layer name
    (setq layerColor (cdr (assoc 62 layerTable)))  ; Get the layer color
    
    ; Check if this color already exists in the association list
    (setq existingEntry (assoc layerColor colorMap))
    
    (if existingEntry
      (progn
        ; If the color exists, get the first layer with that color
        (setq firstLayer (cdr existingEntry))
        
        ; Select all objects on the current layer
        (setq selSet (ssget "X" (list (cons 8 layerName))))  ; Select objects on the current layer
        
        ; If objects are found, change their layer
        (if selSet
          (command "_.CHPROP" selSet "" "_LA" firstLayer)
        )
        
        ; Delete the current layer
        (command "_.LAYER" "_D" layerName)
        
        ; Update the association list to reflect the merged layers
        (setq colorMap (subst (cons layerColor firstLayer) existingEntry colorMap))
      )
      (progn
        ; If color doesn't exist, add the layer to the association list
        (setq colorMap (cons (cons layerColor layerName) colorMap))
      )
    )
    
    ; Move to the next layer
    (setq layerTable (tblnext "layer"))
  )
  
  (princ "\nLayers with the same color have been merged.")
  (princ)
)
