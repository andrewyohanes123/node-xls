/**
 * Cell comments (notes): xl/commentsN.xml plus the legacy VML drawing that
 * Excel uses to position and display the note boxes.
 */
import { CellAddress, Comments } from '../types.js';
import { WriteContext } from '../xlsx/context.js';
import { CONTENT_TYPES, NS, REL_TYPES, XML_HEADER, escapeXml, textElement } from '../xlsx/xml.js';

export interface CommentEntry {
  ref: string;
  addr: CellAddress;
  comments: Comments;
}

/**
 * Writes the comments + VML parts for a sheet and returns the
 * `<legacyDrawing>` element for the worksheet ('' when there are none).
 */
export function writeComments(ctx: WriteContext, entries: CommentEntry[], sheetPart: string): string {
  if (entries.length === 0) return '';
  const { pkg } = ctx;
  const seq = ++ctx.seq.comments;
  // VML shape ids come in blocks of 1024 (o:idmap); blocks must be unique across the workbook
  const blocks = Array.from({ length: Math.ceil(entries.length / 1023) }, () => ++ctx.seq.vmlBlocks);
  const shapeId = (i: number) => blocks[Math.floor(i / 1023)] * 1024 + (i % 1023) + 1;
  const authors: string[] = [];
  const authorId = (a: string) => {
    let i = authors.indexOf(a);
    if (i === -1) i = authors.push(a) - 1;
    return i;
  };

  const commentXml = entries
    .map(({ ref, comments }) => {
      const author = comments[0]?.a ?? '';
      const text =
        comments.length === 1
          ? String(comments[0].t ?? '')
          : comments.map((c) => (c.a ? `${c.a}:\n${c.t ?? ''}` : String(c.t ?? ''))).join('\n');
      // Excel writes note text as a run in its default note font (Tahoma 9).
      return `<comment ref="${ref}" authorId="${authorId(author)}"><text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>${textElement(
        text
      )}</r></text></comment>`;
    })
    .join('');
  const authorsXml = authors.map((a) => `<author>${escapeXml(a)}</author>`).join('');
  const commentsPart = `xl/comments${seq}.xml`;
  pkg.addPart(
    commentsPart,
    `${XML_HEADER}<comments xmlns="${NS.main}"><authors>${authorsXml}</authors><commentList>${commentXml}</commentList></comments>`,
    CONTENT_TYPES.comments
  );
  pkg.addRel(sheetPart, REL_TYPES.comments, commentsPart);

  const shapes = entries
    .map(({ addr, comments }, i) => {
      const visible = comments.hidden === false;
      const { r, c } = addr;
      return (
        `<v:shape id="_x0000_s${shapeId(i)}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${
          i + 1
        };visibility:${visible ? 'visible' : 'hidden'}" fillcolor="#ffffe1" o:insetmode="auto">` +
        '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${r}, 2, ${c + 3}, 15, ${
          r + 4
        }, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column>${visible ? '<x:Visible/>' : ''}</x:ClientData>` +
        '</v:shape>'
      );
    })
    .join('');
  const vml =
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${blocks.join(',')}"/></o:shapelayout>` +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
    '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
    shapes +
    '</xml>';
  const vmlPart = `xl/drawings/vmlDrawing${seq}.vml`;
  pkg.addDefault('vml', CONTENT_TYPES.vml);
  pkg.addPart(vmlPart, vml);
  const rid = pkg.addRel(sheetPart, REL_TYPES.vmlDrawing, vmlPart);
  return `<legacyDrawing r:id="${rid}"/>`;
}
