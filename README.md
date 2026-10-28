# EngSoft Forms — automação universal para Brave

Ferramenta gratuita e editável para auxiliar estudantes a **preencher (não enviar)** seis questões por capítulo em um Google Forms, lendo um único documento Word. Criado para Windows + Brave.

## Instalação (uma vez)

1. Instale o [Brave](https://brave.com/pt-br/download/) e o [Python para Windows](https://www.python.org/downloads/windows/). Durante a instalação do Python, habilite a opção de adicionar o Python ao PATH, se aparecer.
2. Baixe o ZIP deste projeto, extraia todos os arquivos em uma pasta (não execute dentro do ZIP).
3. Dê dois cliques em `INSTALAR.bat` ou abra o CMD nessa pasta e rode `py -3 -m pip install -r requirements.txt`.

## Preenchendo qualquer capítulo

1. Abra `MODELO_UNIVERSAL.docx` no Word e faça uma **cópia** para seu uso pessoal.
2. Preencha nome, matrícula e e-mail; não publique a cópia preenchida.
3. Edite os cabeçalhos `CAPÍTULO 5`, `CAPÍTULO 7`, `CAPÍTULO 8` para os números desejados (ex.: `CAPÍTULO 9 — Manutenção`). Os números devem ser únicos.
4. Em cada cabeçalho, preencha **Tópico no formulário** com o texto EXATO de uma opção disponível na pergunta `Tópico da questão` (bolinha) do Forms. A ferramenta **não pode criar opções novas no formulário**. Se o professor não disponibilizou um capítulo, ele não será marcado; peça a atualização do Forms.
5. Mantenha **seis questões por capítulo**, com enunciado, gabarito, citações e referências e pertinência. Para mais capítulos, copie no Word um bloco inteiro do cabeçalho até a sexta questão; para menos capítulos, exclua blocos inteiros.
6. Salve sua cópia. **Arraste a cópia do Word para `INICIAR.bat`**; ou execute `INICIAR.bat` e edite o modelo padrão antes. O terminal mostrará os capítulos detectados e perguntará qual número preencher.
7. O Brave abre seis abas e tenta preencher o formulário. **Confira o tópico, os dados pessoais e todas as respostas antes de clicar em Enviar.**

## Teste antes de abrir o navegador

Arraste o Word para `VALIDAR.bat` e escolha um capítulo. Isso valida a estrutura do Word, não a veracidade das respostas.

No CMD, também é possível executar:

```bat
py -3 preencher_brave.py "minhas_questoes.docx" --listar
py -3 preencher_brave.py "minhas_questoes.docx" --cap 9 --validar
py -3 preencher_brave.py "minhas_questoes.docx" --cap 9
```

## Limites e cuidados

- O formulário Google Forms é fixado no código (`FORM_URL`); para outra disciplina, altere essa URL e adapte os rótulos de campos em `ALIASES`.
- A ferramenta pode deixar campos sem preencher caso o professor altere o formulário, peça login ou use opções de tópico diferentes das digitadas. Confira todos os campos, especialmente **Tópico da questão**.
- Cada execução preenche **um capítulo, seis abas**. Repita para outro capítulo. Não existe envio automático.
- Não garante correção factual, originalidade das questões, conformidade com o edital ou permissão para automatizar formulários. Revise as respostas e siga as regras do professor.
- Se o Python não for localizado, no CMD tente `py -3 --version`. Caso falhe, reinstale o Python e reabra o terminal.
- Para usar em outros sistemas além do Windows, execute o script Python diretamente; os BATs são específicos do Windows.

## Privacidade

Não envie arquivos pessoais `.docx` preenchidos para repositórios públicos. A automação usa dados locais do Word para preencher o formulário selecionado no navegador.

## Licença

MIT. Consulte `LICENSE`.
