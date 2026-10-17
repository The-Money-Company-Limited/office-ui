import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import type { DocModel } from "@extend-ai/react-docx-doc-model";
import {
  DocxEditorViewer,
  defaultStarterModel,
  useDocxEditor,
  useDocxTrackChanges,
  type DocxTrackedChange,
} from "../../packages/react-viewer/src/editor";

function insertedTextModel(): DocModel {
  return {
    ...defaultStarterModel,
    nodes: [
      {
        type: "paragraph",
        sourceXml:
          '<w:p><w:r><w:t>Before </w:t></w:r><w:ins w:id="17" w:author="Ada"><w:r><w:t>new</w:t></w:r></w:ins></w:p>',
        children: [
          { type: "text", text: "Before " },
          { type: "text", text: "new" },
        ],
      },
    ],
  };
}

describe("annotation editor controls", () => {
  it("exposes revision ids and safe accept/reject actions", () => {
    let changes: DocxTrackedChange[] = [];
    let acceptChange: unknown;
    let rejectChange: unknown;
    const markup = renderToStaticMarkup(
      React.createElement(function TrackedChangesProbe(): React.JSX.Element {
        const editor = useDocxEditor({
          starterModel: insertedTextModel(),
          initialShowTrackedChanges: true,
        });
        const tracked = useDocxTrackChanges(editor);
        changes = tracked.trackedChanges;
        acceptChange = tracked.acceptChange;
        rejectChange = tracked.rejectChange;
        return React.createElement(DocxEditorViewer, { editor });
      })
    );

    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({
      kind: "insertion",
      revisionId: "17",
      location: { kind: "paragraph", nodeIndex: 0 },
    });
    expect(markup).toContain('data-docx-tracked-change="insertion"');
    expect(acceptChange).toBeTypeOf("function");
    expect(rejectChange).toBeTypeOf("function");
  });

  it("strikes through the deletion and underlines the insertion that replaces it", () => {
    const model: DocModel = {
      ...defaultStarterModel,
      nodes: [
        {
          type: "paragraph",
          sourceXml:
            '<w:p><w:r><w:t xml:space="preserve">to a </w:t></w:r><w:del w:id="21" w:author="Ada"><w:r><w:delText>modern ebook format</w:delText></w:r></w:del><w:ins w:id="20" w:author="Ada"><w:r><w:t>current e-book format</w:t></w:r></w:ins><w:r><w:t>, such as</w:t></w:r></w:p>',
          children: [
            { type: "text", text: "to a " },
            { type: "text", text: "current e-book format" },
            { type: "text", text: ", such as" },
          ],
        },
      ],
    };
    const markup = renderToStaticMarkup(
      React.createElement(function ReplacementProbe(): React.JSX.Element {
        const editor = useDocxEditor({
          starterModel: model,
          initialShowTrackedChanges: true,
        });
        return React.createElement(DocxEditorViewer, {
          editor,
          mode: "read-only",
        });
      })
    );

    const span = (kind: string) =>
      new RegExp(
        `<span[^>]*data-docx-tracked-change="${kind}"[^>]*style="([^"]*)"[^>]*>([^<]*)</span>`
      ).exec(markup);
    const deletion = span("deletion");
    const insertion = span("insertion");
    expect(deletion?.[2]).toBe("modern ebook format");
    expect(deletion?.[1]).toContain("text-decoration:line-through");
    expect(insertion?.[2]).toBe("current e-book format");
    expect(insertion?.[1]).toContain("text-decoration:underline");
  });
});
