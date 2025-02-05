# SleekDoc-Converter ✨
Este repositório contém um código Python que converte documentos no formato `.docx` para `.pdf`, preservando a formatação de texto, imagens e estilos. Todo o código foi criado manualmente e está comentado detalhadamente para auxiliar no aprendizado e compreensão da linguagem Python.

---

## 🚀 Tecnologias Utilizadas

- **python-docx**: Manipulação de arquivos DOCX.
- **reportlab**: Geração de PDFs.
- **Pillow (PIL)**: Manipulação de imagens.
- **tkinter**: Interface para seleção de arquivos.
- **xml.etree.ElementTree**: Manipulação de XML presente nos arquivos DOCX.

---

## 🛠 Estrutura e Lógica do Código

### 1. Importação de Bibliotecas
O código inicia importando diversas bibliotecas para lidar com:
- Arquivos `.docx`
- Conversão para `.pdf`
- Manipulação de imagens
- Interação com o usuário

---

### 2. Seleção do Arquivo DOCX
- Utiliza `filedialog.askopenfilename()` (tkinter) para que o usuário selecione um arquivo `.docx`.
- Se o arquivo selecionado não for `.docx`, o código solicita a seleção de um novo arquivo.

---

### 3. Extração de Informações do Documento
- **Textos e Formatação**: Extrai parágrafos e propriedades como alinhamento, fonte, tamanho, espaçamento, cores, negrito, itálico, sublinhado, etc.
- **Imagens**: Coleta informações sobre as imagens incorporadas.
- **Margens**: Obtém as margens definidas no documento para aplicação no PDF final.

---

### 4. Conversão de Tamanhos com Regra de Três
- **Conversão de Unidades**:
  - 1 polegada = 1440 twips.
  - 1 polegada = 72 pt.
  - Assim, 1 twip equivale a 1/20 de pt.
- **Cálculo Importante**:
  - O espaçamento de parágrafos é convertido dividindo por **12700** (valor obtido por meio de uma regra de três) e depois ajustado:
    ```python
    padding_after = (padding_after / 12700) + 10 + 5
    padding_before = (padding_before / 12700) + 10 + 5
    ```
  - Garante que o espaçamento entre parágrafos no PDF seja coerente com o original do DOCX.

---

### 5. Extração de Imagens e Conversão para JPG
- Percorre a estrutura XML do DOCX para localizar as imagens incorporadas.
- Extrai os dados binários das imagens e os salva como arquivos `.jpg` na pasta `imagens/`.

> :warning: **Limitação Importante**  
> Se houver espaços em branco, quebras de linha sem conteúdo ou textos vazios após as imagens no DOCX, esses elementos serão interpretados como imagens, podendo causar erros no sistema.

---

### 6. Geração do PDF Final
- **Texto**: Cada parágrafo do DOCX é convertido para um objeto `Paragraph` no PDF.
- **Imagens**: Inseridas no PDF de acordo com as dimensões extraídas.
- **Margens e Estilos**: Margens originais e estilizações (negrito, itálico, alinhamento, cores) são preservadas.
- O resultado é um PDF que reflete fielmente o documento original.

---

## 📚 Como Usar

1. **Clone este repositório:**
   ```bash
   git clone https://github.com/seuusuario/seurepositorio.git
   ```
2. **Instale as dependências:**
```bash
pip install python-docx reportlab pillow
```
3. **Execute o script:**
```bash
python index.py
```
4. **Selecione um arquivo .docx quando solicitado.**
O PDF gerado será salvo na pasta ```convertidos/.```

---

## 🔍 Observações Importantes

- **Uso de Matemática**:  
  Foram utilizadas fórmulas matemáticas para converter tamanhos, como a conversão de twips para pt utilizando o fator **12700**, obtido por meio de uma regra de três.

- **Criação Manual**:  
  Todo o código foi desenvolvido manualmente e está amplamente comentado para facilitar o entendimento de cada função e conceito da linguagem Python.

- **Limitações**:  
  :warning: Caso haja espaços em branco, quebras de linha sem conteúdo ou textos vazios após as imagens no DOCX, esses elementos serão contabilizados como imagens, o que pode gerar erros no sistema.

---

## 👤 Autor

Este código foi criado para estudo e aprimoramento na linguagem Python. Cada trecho foi comentado detalhadamente para ajudar na compreensão dos conceitos e na assimilação do código.

---

## 📄 Licença

Este projeto é de uso livre. Sinta-se à vontade para modificar e aprimorar!

