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
        p("O trabalho transformou uma ideia em linguagem natural — um site institucional para uma consultoria de governança de dados e IA — em um protótipo visual e, depois, em uma página web funcional, usando a IA (Claude) como par de desenvolvimento do início ao fim, sem escrita manual de código. Esse é o núcleo do \"vibe coding\": comunica-se a intenção, e o papel humano se desloca de quem escreve cada linha para quem define o objetivo, avalia o resultado e corrige o rumo. O processo teve três camadas revisadas em sequência: um esboço estrutural de baixa fidelidade, um protótipo visual de alta fidelidade e, por fim, a página HTML/CSS real, navegável e testável localmente."),

        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 300, after: 160 },
          children: [new TextRun({ text: "Como a empresa poderia usar Vibe Coding" })],
        }),
        p("Essa não é uma possibilidade hipotética: é uma experiência que já vivemos na Data Hawk Solution, com dois papéis se complementando."),
        p("Minha sócia implantou o site institucional da Data Hawk inteiramente pelo assistente de IA, sem nenhum conhecimento técnico de desenvolvimento, chegando a uma versão muito próxima da final e implementando até os serviços comercializáveis online, incluindo formas de pagamento — evidência de como o vibe coding reduz a barreira entre ter uma ideia de negócio e colocá-la de pé."),
        p("Em seguida, assumi o desenvolvimento dos produtos e evoluí o nosso primeiro produto para uma versão que usa agentes de IA na geração de conteúdo em tempo real. Por ter mais conhecimento técnico, implementei também componentes de arquitetura mais profissionais, agregando flexibilidade na implantação em cada cliente — decisões de engenharia que vão além do que o vibe coding resolve sozinho."),
        p("Outro ponto relevante foi a criação de um repositório no Git, que centralizou o código e organizou o trabalho em equipe, com histórico de mudanças e um ponto único de verdade sobre o que está em produção."),
        p("Essa combinação resume o potencial do vibe coding para uma consultoria pequena como a Data Hawk: ele abre a porta para qualquer pessoa da equipe colocar uma ideia em produção rapidamente, enquanto a experiência técnica garante a robustez, a flexibilidade e a governança que um cliente regulado exige."),

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
