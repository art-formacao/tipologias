# Aplicação de Tipologias TJA

## Estrutura de cada perfil

Crie as pastas com esta organização:

```text
Tipo de Reboque/
  Lona/
    Lonas standard/
      Fotos/
        1.jpg
        2.jpg
        3.jpg
      Equipamentos Associados/
        2 Réguas de travamento.jpg
      Especificações Técnicas e Descrição.xlsx
```

Para rígidos com níveis intermédios pode usar, por exemplo:

```text
Tipo de Rígido/Frigorífico/Bi Temperatura 1 eixo/9t/
```

Dentro de `9t` ficam as pastas `Fotos`, `Equipamentos Associados` e o ficheiro Excel.

## Regras automáticas

- As imagens de `Fotos` aparecem no perfil por ordem natural do nome: `1`, `2`, `3`, `4`…
- O nome de cada ficheiro em `Equipamentos Associados` é apresentado como nome do equipamento. A extensão da imagem é removida.
- São aceites imagens JPG, JPEG, PNG, WEBP, GIF e AVIF.
- No Excel, a coluna A contém o nome da especificação e a coluna B contém o valor.
- A linha `Descrição` é usada como descrição do perfil e não aparece na grelha técnica.
- O Excel deve chamar-se exatamente `Especificações Técnicas e Descrição.xlsx`.

## Publicação

Copie todo o conteúdo deste pacote para a raiz do repositório. Em **Settings → Pages**, escolha **GitHub Actions** como origem. A cada alteração enviada para o ramo `main`, o processo incluído volta a ler as pastas e publica a aplicação atualizada.

Não edite manualmente `data-manifest.json`: este ficheiro é recriado automaticamente durante a publicação.
