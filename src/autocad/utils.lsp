;; Define uma função chamada ReplaceString com três parâmetros:
;; - searchText: o texto que será procurado dentro da string de conteúdo.
;; - replaceText: o texto que substituirá o texto encontrado.
;; - contentString: a string onde será feita a busca e substituição.
;; A barra "/" indica que não há variáveis locais declaradas além dos argumentos.
(defun ReplaceString (searchText replaceText contentString /)

  ;; Verifica se a string de busca (searchText) existe dentro da string principal (contentString).
  ;; A função vl-string-search retorna a posição do texto encontrado ou NIL se não encontrar.
  (if (vl-string-search searchText contentString)

    ;; Se o texto for encontrado, a função vl-string-subst realiza a substituição:
    ;; substitui a primeira ocorrência de searchText por replaceText dentro de contentString.
    (vl-string-subst replaceText searchText contentString)

    ;; Caso contrário, se searchText não for encontrado, retorna a string original (sem alterações).
    contentString
  )
)

;; Define uma função chamada sort-by-length que recebe uma lista lst como argumento.
(defun sort-by-length (lst)
  
  ;; Utiliza a função vl-sort do AutoLISP para ordenar a lista lst.
  ;; A ordenação será feita com base em uma função de comparação personalizada.
  (vl-sort lst

    ;; Define a função de comparação que será usada para ordenar os elementos.
    ;; Esta função compara dois elementos da lista (a e b).
    (function
      (lambda (a b)

        ;; Compara o comprimento (número de caracteres) da primeira string de cada sublista.
        ;; (car a) e (car b) acessam o primeiro elemento (a string) de cada sublista.
        ;; A função strlen retorna o comprimento da string.
        ;; A ordenação será crescente: sublistas com strings menores vêm antes.
        (< (strlen (car a)) (strlen (car b)))
      )
    )
  )
)

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