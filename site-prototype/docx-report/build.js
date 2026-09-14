const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun,
  AlignmentType, PageBreak
} = require("docx");

const IMG_DIR = "/home/user/datagov-builder/site-prototype/images/";

const bodyStyle = { size: 22 }; // 11pt

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 200, line: 300 },
    children: [new TextRun({ text, ...bodyStyle, ...opts })],
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 360 },
    children: [new TextRun({ text, italics: true, size: 20, color: "555555" })],
  });
}

function figureTitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 300, after: 120 },
    children: [new TextRun({ text, bold: true, size: 22 })],
  });
}

const doc = new Document({
  sections: [
    {
      properties: {},
      children: [
        new Paragraph({
          heading: HeadingLevel.TITLE,
          spacing: { after: 120 },
          children: [new TextRun({ text: "Vibe Coding aplicado a um protótipo institucional" })],
        }),
        new Paragraph({
          spacing: { after: 300 },
          children: [new TextRun({ text: "Estudo de caso: protótipo de site para a Data Hawk Solution", italics: true, size: 22, color: "555555" })],
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 160 },
          children: [new TextRun({ text: "Resumo do processo" })],
        }),
        p("O trabalho consistiu em transformar uma ideia descrita em linguagem natural — um site institucional para uma consultoria de governança de dados e IA — em um protótipo visual e, em seguida, em uma página web funcional, usando um assistente de IA (Claude) como par de desenvolvimento ao longo de todo o processo. Não houve escrita manual de código em nenhuma etapa: cada resultado foi produzido a partir de instruções em português descrevendo a intenção (\"quero uma página inicial, com apresentação da consultoria, e subpáginas de serviços e contato, com interface limpa e moderna\"), refinadas por rodadas curtas de revisão e ajuste."),
        p("Esse é o núcleo do que se chama de \"vibe coding\": em vez de especificar sintaxe, estruturas de dados ou arquitetura de antemão, o desenvolvedor comunica a intenção e o resultado desejado, e deixa que o modelo gere o artefato (design, texto, código), revisando e redirecionando pelo resultado obtido, não pelo caminho percorrido. O papel humano não desaparece — desloca-se: de quem escreve cada linha para quem define objetivo, avalia o resultado, corrige rumo e garante que o conteúdo gerado seja verdadeiro e coerente com a realidade da empresa, e não apenas texto genérico de marketing."),
        p("Na prática, o processo teve três camadas: (1) um esboço estrutural de baixa fidelidade da ideia original, para validar a organização do conteúdo antes de investir em visual; (2) um protótipo visual de alta fidelidade (tipografia, cores, ícones, hierarquia), construído em uma ferramenta de design; e (3) a tradução desse protótipo em uma página HTML/CSS real, navegável e testável localmente. Cada camada foi revisada antes de avançar para a próxima — inclusive com uma etapa de conferência automatizada do conteúdo gerado, para pegar inconsistências de texto e de layout antes da entrega."),

        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 300, after: 160 },
          children: [new TextRun({ text: "Como a empresa poderia usar Vibe Coding" })],
        }),
        p("Essa não é uma possibilidade hipotética: é uma experiência que já estamos vivendo na própria Data Hawk Solution, com dois papéis diferentes se complementando."),
        p("Minha sócia implantou o site institucional da Data Hawk inteiramente através do assistente de IA, sem ter absolutamente nenhum conhecimento técnico de desenvolvimento. Nessa implantação, ela já conseguiu colocar no ar uma versão muito próxima da versão visual final da consultoria — e foi além, implementando também os serviços que a empresa já poderia comercializar online, incluindo formas de pagamento. Isso evidencia o quanto o vibe coding reduz a barreira entre ter uma ideia de negócio e colocá-la de pé, mesmo sem formação técnica."),
        p("Em seguida, assumi o desenvolvimento dos produtos e evoluí o nosso primeiro produto para uma versão que utiliza agentes de IA na geração de conteúdo em tempo real, em vez de um fluxo estático. Por ter mais conhecimento técnico, também implementei componentes de arquitetura mais profissionais, agregando flexibilidade no momento da implantação em cada cliente — decisões de engenharia (modularidade, configuração por ambiente, isolamento de dados sensíveis) que vão além do que o vibe coding resolve sozinho."),
        p("Outro ponto relevante desse processo foi a criação de um repositório no Git, que permitiu centralizar o código e passar a trabalhar em equipe de forma organizada, com histórico de mudanças e um ponto único de verdade sobre o que está em produção, em vez de cada pessoa guardando sua própria versão do projeto."),
        p("Essa combinação resume bem o potencial do vibe coding para uma consultoria pequena como a Data Hawk Solution: ele abre a porta para que qualquer pessoa da equipe — técnica ou não — coloque uma ideia de negócio em produção rapidamente, enquanto a experiência técnica continua sendo o que garante que essa ideia evolua com a robustez, a flexibilidade e a governança que um cliente regulado exige."),

        new Paragraph({
          children: [new PageBreak()],
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 0, after: 160 },
          children: [new TextRun({ text: "Imagens" })],
        }),

        figureTitle("a) Modelo da ideia original"),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              type: "png",
              data: fs.readFileSync(IMG_DIR + "01-ideia-original-wireframe.png"),
              transformation: { width: 460, height: 556 },
            }),
          ],
        }),
        caption("Figura 1 — Rascunho estrutural de baixa fidelidade (wireframe) da página inicial, usado para validar a organização do conteúdo antes do design visual."),

        new Paragraph({
          children: [new PageBreak()],
        }),

        figureTitle("b) Página web HTML criada"),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              type: "png",
              data: fs.readFileSync(IMG_DIR + "02-pagina-final-index.png"),
              transformation: { width: 380, height: 805 },
            }),
          ],
        }),
        caption("Figura 2 — Página inicial (index.html) do protótipo final, em HTML e CSS, navegável e testável localmente no navegador."),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(
    "/home/user/datagov-builder/site-prototype/docx-report/relatorio-vibe-coding.docx",
    buffer
  );
  console.log("done");
});
