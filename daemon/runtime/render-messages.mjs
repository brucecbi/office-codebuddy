// Canonical message upserts reconcile incomplete/divergent deltas by block index.
// Unknown-id deltas are buffered until a complete message arrives. No thinking.
export function createMessageRenderer(emit, turnId) {
  const messages = new Map();
  const current = new Map();
  let serial = 0;
  function record(parent, id) {
    const key = parent + ":" + id;
    if (!messages.has(key))
      messages.set(key, { id: `${turnId}:${key}`, blocks: new Map(), done: false });
    return messages.get(key);
  }
  function publish(r) {
    const text = [...r.blocks]
      .sort((a, b) => a[0] - b[0])
      .map(([, t]) => t)
      .join("");
    if (text !== r.last && (text || r.last !== undefined)) {
      r.last = text;
      emit({ event: "assistant_message", message_id: r.id, text, turn_id: turnId });
    }
  }
  return (msg) => {
    const parent = msg.parent_tool_use_id || "main";
    if (msg.type === "stream_event") {
      const e = msg.event || {};
      if (e.type === "message_start" && e.message?.id) current.set(parent, e.message.id);
      const id = msg._messageId || e.message?.id || current.get(parent);
      if (!id) return;
      const r = record(parent, id);
      if (r.done) return;
      const index = e.index ?? 0;
      if (e.type === "content_block_start" && e.content_block?.type === "text") {
        r.blocks.set(index, e.content_block.text || "");
      } else if (e.delta?.type === "text_delta") {
        r.blocks.set(index, (r.blocks.get(index) || "") + (e.delta.text || ""));
      } else return;
      publish(r);
    } else if (msg.type === "assistant") {
      const m = msg.message ?? msg;
      const blocks =
        typeof m.content === "string" ? [{ type: "text", text: m.content }] : m.content || [];
      const id = m.id || msg.uuid || current.get(parent) || `complete-${++serial}`;
      const r = record(parent, id);
      r.blocks.clear();
      blocks.forEach((b, index) => {
        if ((b.type === "text" || b.type === "output_text") && typeof b.text === "string")
          r.blocks.set(index, b.text);
        if (b.type === "tool_use")
          emit({
            event: "tool_use_announce",
            tool_id: b.id,
            tool: b.name,
            input: b.input,
            parent_tool_use_id: msg.parent_tool_use_id,
            turn_id: turnId,
          });
      });
      r.done = true;
      publish(r);
      current.delete(parent);
    } else if (msg.type === "user") {
      for (const b of msg.message?.content ?? msg.content ?? []) {
        if (b.type === "tool_result")
          emit({
            event: "tool_result_announce",
            tool_id: b.tool_use_id,
            output: b.content,
            status: b.is_error ? "error" : "completed",
            turn_id: turnId,
          });
      }
    }
  };
}
