(defun CreateLayerIfNotExists (layerName colorIndex / acadObj doc layers layer)
  ;; Define uma função para criar uma layer com nome e cor, caso ainda não exista
  ;; A barra (/) indica que as variáveis listadas a seguir são locais à função.
  ;; 1 = Vermelho
  ;; 2 = Amarelo
  ;; 3 = Verde
  ;; 4 = Ciano
  ;; 5 = Azul
  ;; 6 = Magenta
  ;; 7 = Branco/Preto (dependendo do fundo)

  (setq acadObj (vlax-get-acad-object)) ; Obtém o AutoCAD
  (setq doc (vla-get-ActiveDocument acadObj)) ; Documento ativo
  (setq layers (vla-get-Layers doc)) ; Coleção de layers

  ;; Verifica se a layer já existe
  (if (not (tblsearch "LAYER" layerName))
    (progn
      ;; Cria a nova layer
      (setq layer (vla-Add layers layerName))
      ;; Define a cor da layer
      (vla-put-Color layer colorIndex)
      ;; Mensagem de confirmação
      (princ (strcat "\nLayer criada: " layerName " com cor " (itoa colorIndex)))
    )
    ;; Caso já exista
    (princ (strcat "\nLayer já existe: " layerName))
  )
)
(defun c:OrganizarPorTextoEmLayers (/ ss idx obj entity substitutions contentString searchText att attColl)
  
  ;; Lista de palavras-chave a procurar nos textos
  ;; Cada palavra corresponde ao nome da layer para onde o objeto será movido
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
      (setq idx 0) ; índice de controle do loop

      ;; Loop sobre todos os objetos selecionados
      (while (< idx (sslength ss))
        (setq entity (ssname ss idx)) ; Pega a entidade na posição idx

        (if entity
          (progn
            (setq obj (vlax-ename->vla-object entity)) ; Converte para objeto VLA para acessar propriedades

            ;; Verifica o tipo de objeto e extrai o conteúdo do texto
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
                        (setq contentString (vla-get-TextString att)) ; Conteúdo do atributo

                        (if contentString
                          (progn
                            ;; Verifica se o conteúdo do atributo contém uma das palavras-chave
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
                    ;; O conteúdo do bloco já foi processado via atributos, então limpa a variável
                    (setq contentString nil)
                  )
                )
              )
              ;; Se não for um dos tipos acima, ignora
              (T (setq contentString nil))
            )

            ;; Se for um texto (e não um bloco) e conteúdo foi encontrado
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
        ;; Próxima entidade
        (setq idx (1+ idx))
      )

      ;; Mensagem de sucesso
      (princ "\nSubstituições concluídas.")
    )
    ;; Nenhum objeto válido foi encontrado
    (princ "\nNenhum objeto do tipo TEXT, MTEXT, MULTILEADER ou INSERT encontrado.")
  )

  (princ) ; Finaliza de forma limpa
)
(c:OrganizarPorTextoEmLayers)

