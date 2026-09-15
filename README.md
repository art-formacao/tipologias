# Aplicação de Tipologias TJA

## Localização das duas pastas

As pastas devem ficar separadas na raiz do repositório, ao lado de `index.html`:

```text
index.html
Tipo de Reboque/
Tipo de Rígido/
```

Quando o GitHub mostra `Tipo de Reboque/Lona/Lonas standard` numa única linha e apresenta “This path skips through empty directories”, está apenas a compactar visualmente pastas com um único caminho. As três pastas continuam separadas e o caminho está correto.

## Conteúdo de cada perfil

```text
Tipo de Reboque/Lona/Lonas standard/
  Fotos/
    1.jpg
    2.jpg
  Equipamentos Associados/
    2 Réguas de travamento.jpg
  Especificações Técnicas e Descrição.xlsx
```

As fotografias são ordenadas pelo número do nome. Nos equipamentos, o nome do ficheiro torna-se o nome apresentado. No Excel, a coluna A contém o nome do campo e a coluna B o valor. A linha `Descrição` alimenta a descrição do perfil.

## Ficheiro de publicação obrigatório

O ficheiro de publicação não pode ficar na raiz. O caminho obrigatório é:

```text
.github/workflows/deploy-pages.yml
```

O ficheiro `generate-data-manifest.mjs` fica agora na raiz para simplificar o carregamento.

Em **Settings → Pages**, escolha **GitHub Actions** como origem. Depois de enviar alterações para `main`, confirme no separador **Actions** que “Publicar aplicação no GitHub Pages” terminou com sucesso.
