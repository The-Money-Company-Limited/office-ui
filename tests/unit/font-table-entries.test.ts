import { describe, expect, it } from "vitest";
import { fontTableEntries } from "../../packages/react-viewer/src/font-table";

describe("fontTableEntries", () => {
  it("keeps a self-closing font from borrowing the next font's embedded faces", () => {
    // Google Docs exports declare system fonts as empty elements and embed only
    // the custom face, e.g. Arial followed by an embedded Happy Monkey.
    const xml =
      '<w:fonts xmlns:w="w" xmlns:r="r"><w:font w:name="Arial"/>' +
      '<w:font w:name="Happy Monkey"><w:embedRegular w:fontKey="{00000000-0000-0000-0000-000000000000}" r:id="rId1" w:subsetted="0"/></w:font>' +
      "</w:fonts>";
    expect(fontTableEntries(xml)).toEqual([
      {
        family: "Happy Monkey",
        xml: '<w:embedRegular w:fontKey="{00000000-0000-0000-0000-000000000000}" r:id="rId1" w:subsetted="0"/>',
      },
    ]);
  });

  it("returns each font with its own children", () => {
    const xml =
      '<w:fonts><w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>' +
      '<w:font w:name="Trebuchet MS" ><w:embedBold r:id="rId2"/></w:font>' +
      '<w:font w:name="Symbol" /></w:fonts>';
    expect(fontTableEntries(xml).map((entry) => [entry.family, entry.xml])).toEqual([
      ["Calibri", '<w:panose1 w:val="020F0502020204030204"/>'],
      ["Trebuchet MS", '<w:embedBold r:id="rId2"/>'],
    ]);
  });
});
