(defun CreateLayerIfNotExists (layerName colorIndex / acadObj doc layers layer)
  ;; Define uma fun��o para criar uma layer com nome e cor, caso ainda n�o exista
  ;; A barra (/) indica que as vari�veis listadas a seguir s�o locais � fun��o.
  ;; 1 = Vermelho
  ;; 2 = Amarelo
  ;; 3 = Verde
  ;; 4 = Ciano
  ;; 5 = Azul
  ;; 6 = Magenta
  ;; 7 = Branco/Preto (dependendo do fundo)

  (setq acadObj (vlax-get-acad-object)) ; Obt�m o AutoCAD
  (setq doc (vla-get-ActiveDocument acadObj)) ; Documento ativo
  (setq layers (vla-get-Layers doc)) ; Cole��o de layers

  ;; Verifica se a layer j� existe
  (if (not (tblsearch "LAYER" layerName))
    (progn
      ;; Cria a nova layer
      (setq layer (vla-Add layers layerName))
      ;; Define a cor da layer
      (vla-put-Color layer colorIndex)
      ;; Mensagem de confirma��o
      (princ (strcat "\nLayer criada: " layerName " com cor " (itoa colorIndex)))
    )
    ;; Caso j� exista
    (princ (strcat "\nLayer j� existe: " layerName))
  )
)
(defun c:OrganizarPorTextoEmLayers (/ ss idx obj entity substitutions contentString searchText att attColl)
  
  ;; Lista de palavras-chave a procurar nos textos
  ;; Cada palavra corresponde ao nome da layer para onde o objeto ser� movido
  (setq substitutions '(
    ("FV" 8)
    ("FO" 8)
	("PIT" 8)
	("PCV" 8)
	("LE" 8)
	("LIT" 8)
	("DIT" 8)
	("FIT" 8)
	("FE" 8)
	("PIA" 8)
	("TIA" 8)
	("SIC" 8)
  ))
  
  (foreach pair substitutions
    (setq layerName (car pair))
    (setq layerColor (cadr pair))
    (CreateLayerIfNotExists layerName layerColor) ; Garante que a layer exista
  )

  ;; Seleciona todos os objetos do tipo TEXT, MTEXT, MULTILEADER ou INSERT (blocos)
  (setq ss (ssget "X" '((0 . "TEXT,MTEXT,MULTILEADER,INSERT"))))

  (if ss
    (progn
      (setq idx 0) ; �ndice de controle do loop

      ;; Loop sobre todos os objetos selecionados
      (while (< idx (sslength ss))
        (setq entity (ssname ss idx)) ; Pega a entidade na posi��o idx

        (if entity
          (progn
            (setq obj (vlax-ename->vla-object entity)) ; Converte para objeto VLA para acessar propriedades

            ;; Verifica o tipo de objeto e extrai o conte�do do texto
            (cond
              ((= (vla-get-ObjectName obj) "AcDbMLeader") ; MLeader
                (setq contentString (vla-get-TextString obj))
              )
              ((= (vla-get-ObjectName obj) "AcDbMText") ; MText
                (setq contentString (vla-get-TextString obj))
              )
              ((= (vla-get-ObjectName obj) "AcDbText") ; Texto simples
                (setq contentString (vla-get-TextString obj))
              )
              ((= (vla-get-ObjectName obj) "AcDbBlockReference") ; Bloco
                ;; Se o bloco tem atributos
                (if (vlax-property-available-p obj 'HasAttributes)
                  (progn
                    (setq attColl (vlax-invoke obj 'GetAttributes)) ; Pega os atributos do bloco

                    ;; Loop pelos atributos
                    (if attColl
                      (foreach att attColl
                        (setq contentString (vla-get-TextString att)) ; Conte�do do atributo

                        (if contentString
                          (progn
                            ;; Verifica se o conte�do do atributo cont�m uma das palavras-chave
                            (foreach pair substitutions
							  (setq layerName (car pair))
                              (if (vl-string-search layerName contentString)
                                (vla-put-Layer obj layerName) ; Move o bloco para a layer correspondente
                              )
                            )
                          )
                        )
                      )
                    )
                    ;; O conte�do do bloco j� foi processado via atributos, ent�o limpa a vari�vel
                    (setq contentString nil)
                  )
                )
              )
              ;; Se n�o for um dos tipos acima, ignora
              (T (setq contentString nil))
            )

            ;; Se for um texto (e n�o um bloco) e conte�do foi encontrado
            (if (and contentString
                     (not (= (vla-get-ObjectName obj) "AcDbBlockReference")))
              (progn
                (foreach pair substitutions
				  (setq layerName (car pair))
                  (if (vl-string-search layerName contentString)
                    (vla-put-Layer obj layerName) ; Move para a layer correspondente
                  )
                )
              )
            )
          )
        )
        ;; Pr�xima entidade
        (setq idx (1+ idx))
      )

      ;; Mensagem de sucesso
      (princ "\nSubstitui��es conclu�das.")
    )
    ;; Nenhum objeto v�lido foi encontrado
    (princ "\nNenhum objeto do tipo TEXT, MTEXT, MULTILEADER ou INSERT encontrado.")
  )

  (princ) ; Finaliza de forma limpa
)
(c:OrganizarPorTextoEmLayers)

