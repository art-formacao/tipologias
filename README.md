# Aplicação TJA — dados gerados pelas pastas

A única fonte de dados é a pasta `Tipologias`. Não existem tipos, categorias ou perfis fixos no código.

## Estrutura

```text
Tipologias/
  Tipo de Reboque/
    Lona/
      Lona standard/
        Fotos/
          1.jpg
          2.jpg
          3.jpg
        Equipamentos Associados/
          2 Réguas de travamento.jpg
        Especificações Técnicas e Descrição.xlsx
  Tipo de Rígido/
    Frigorífico/
      Bi Temperatura 1 eixo/
        9t/
          Fotos/
          Equipamentos Associados/
          Especificações Técnicas e Descrição.xlsx
```

- Cada pasta diretamente dentro de `Tipologias` cria um separador principal na aplicação. Se amanhã adicionar outro tipo, este aparece automaticamente.
- As pastas seguintes criam as categorias, como `Lona`, `Cisterna` ou `Frigorífico`.
- Uma pasta é reconhecida como perfil quando contém `Fotos`, `Equipamentos Associados` ou o Excel.
- As fotografias são apresentadas pela ordem numérica dos nomes: 1, 2, 3, 4…
- O nome da fotografia em `Equipamentos Associados` torna-se o nome apresentado no perfil.
- No Excel, a coluna A contém o nome da especificação e a coluna B contém o valor. A linha `Descrição` fornece a descrição do perfil.
- Se uma pasta não existir, não aparece na aplicação.

## GitHub

Coloque `Tipo de Reboque` e `Tipo de Rígido` dentro de `Tipologias`. O ficheiro de publicação deve permanecer em `.github/workflows/deploy-pages.yml`. Em **Settings → Pages**, selecione **GitHub Actions**. Cada envio para `main` volta a ler toda a estrutura e publica os dados atuais.
