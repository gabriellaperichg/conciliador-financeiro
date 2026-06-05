===============================================================
  CONCILIADOR BANCO x CORPORATIVO — Wave 1
  Instruções de uso
===============================================================

PRÉ-REQUISITO (fazer uma vez só)
---------------------------------
1. Instale o Python em python.org/downloads
   → Na instalação, marque "Add Python to PATH"

2. Abra o terminal (Windows: Prompt de Comando ou PowerShell)
   e navegue até a pasta do projeto:
   cd C:\caminho\da\pasta\conciliador

3. Instale as dependências (fazer uma vez só):
   pip install -r requirements.txt
   pip install pyyaml


COMO USAR TODO MÊS
-------------------
1. Copie os dois arquivos Excel para dentro desta pasta:
   - Base do corporativo (ex: Base_Corporativo.xlsx)
   - Extrato do banco    (ex: Base_BTG_-_Londrina.xlsx)

2. Abra o config.yaml e ajuste:
   - unidade: "Londrina"   (ou "Maringá" etc.)
   - banco:   "BTG"        (ou "Itaú")
   - mes:     5            (número do mês)
   - arquivo_corporativo e arquivo_banco  (nome exato dos arquivos)
   - saida:   nome do arquivo que será gerado

3. Rode no terminal:
   python conciliador.py

4. Abra o arquivo Excel gerado e revise:
   → Aba "A Reportar"      → enviar ao corporativo
   → Aba "Divergência"     → conferir valor correto
   → Aba "Pendente"        → checar se foi pago por outro meio
   → Aba "Conciliados"     → auditoria (tudo que bateu)


ESTRUTURA DE ARQUIVOS
----------------------
conciliador/
├── conciliador.py          ← não edite
├── config.yaml             ← edite aqui todo mês
├── requirements.txt        ← não edite
├── LEIAME.txt              ← este arquivo
└── [arquivos Excel aqui]


DÚVIDAS FREQUENTES
-------------------
"O terminal diz que Python não foi encontrado"
→ Reinstale o Python marcando "Add Python to PATH"

"Erro: arquivo não encontrado"
→ Confirme que o nome no config.yaml está igual ao nome
  do arquivo na pasta (incluindo letras maiúsculas)

"Resultado parece errado — muitos pendentes"
→ Confira se unidade e banco estão escritos no config
  exatamente como aparecem na base do corporativo
===============================================================