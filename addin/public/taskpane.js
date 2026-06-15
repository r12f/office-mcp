(() => {
  const ADDIN_VERSION = '0.1.5';
  const PROTOCOL_VERSION = '1.0';
  const AVAILABLE_TOOLS = [
    'word.get_text',
    'word.get_outline',
    'word.get_paragraph',
    'word.find_text',
    'word.get_selection',
    'word.insert_paragraph',
    'word.insert_heading',
    'word.insert_image',
    'word.insert_table',
    'word.insert_page_break',
    'word.insert_list',
    'word.replace_text',
    'word.update_paragraph',
    'word.delete_range',
    'word.apply_formatting',
    'word.read_table',
    'word.update_cell',
    'word.add_row',
    'word.add_column',
    'word.format_cell',
    'word.set_heading_level',
    'word.apply_style',
    'word.add_comment',
    'word.resolve_comment',
    'word.accept_change',
    'word.reject_change',
    'word.save'
  ];

  let socket;
  let instanceId = sessionStorage.getItem('office-mcp.instance-id') || crypto.randomUUID();
  let sessionId = sessionStorage.getItem('office-mcp.session-id') || crypto.randomUUID();
  const cancelledRequests = new Set();
  sessionStorage.setItem('office-mcp.instance-id', instanceId);
  sessionStorage.setItem('office-mcp.session-id', sessionId);

  const statusEl = document.getElementById('status');
  const sessionEl = document.getElementById('session');
  const daemonEl = document.getElementById('daemon');

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Word) {
      setStatus('Unsupported host');
      return;
    }
    if (Office.context.platform !== Office.PlatformType.PC) {
      setStatus('Unsupported platform');
      return;
    }
    setStatus('Connecting');
    connect();
  });

  function connect() {
    const endpoint = `${location.origin.replace('https://', 'wss://')}/addin`;
    daemonEl.textContent = endpoint;
    socket = new WebSocket(endpoint);
    socket.addEventListener('open', () => register());
    socket.addEventListener('message', (event) => handleMessage(event.data));
    socket.addEventListener('close', () => {
      console.warn('office-mcp websocket closed');
      setStatus('Disconnected; retrying');
      setTimeout(connect, 1000 + Math.floor(Math.random() * 1500));
    });
    socket.addEventListener('error', () => setStatus('Connection error'));
  }

  function register() {
    setStatus('Registering');
    send({
      jsonrpc: '2.0',
      id: crypto.randomUUID(),
      method: 'register',
      params: {
        instance_id: instanceId,
        host: {
          app: 'word',
          version: Office.context.diagnostics?.version || null,
          platform: String(Office.context.platform || 'unknown').toLowerCase(),
          build: Office.context.diagnostics?.host || 'Desktop'
        },
        add_in: {
          version: ADDIN_VERSION,
          protocol_version: PROTOCOL_VERSION,
          requirement_sets: probeRequirementSets(),
          supported_features: ['doc.read', 'doc.write']
        }
      }
    });
    announceSession();
  }

  async function announceSession() {
    const document = await getDocumentInfo();
    await enableDocumentAutoOpen();
    console.log('office-mcp session.added', { sessionId, document });
    send({
      jsonrpc: '2.0',
      method: 'session.added',
      params: {
        session_id: sessionId,
        instance_id: instanceId,
        document,
        available_tools: AVAILABLE_TOOLS,
        is_active: null
      }
    });
    sessionEl.textContent = sessionId;
    setStatus('Connected');
  }

  function enableDocumentAutoOpen() {
    return new Promise((resolve) => {
      try {
        Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
        Office.context.document.settings.saveAsync(() => resolve());
      } catch {
        resolve();
      }
    });
  }

  function handleMessage(raw) {
    let message;
    try {
      message = JSON.parse(raw);
    } catch {
      return;
    }
    if (message.method === 'tool.invoke') {
      invokeTool(message).catch((error) => reply(message.id, {
        ok: false,
        error: mapError(error, message.params?.tool),
        elapsed_ms: 0
      }));
    } else if (message.method === 'ping') {
      reply(message.id, { ts: new Date().toISOString() });
    } else if (message.method === 'tool.cancel' && message.params?.request_id) {
      cancelledRequests.add(message.params.request_id);
    }
  }

  async function invokeTool(message) {
    const started = performance.now();
    const { tool, args } = message.params;
    let data;
    if (cancelledRequests.has(message.id)) throw cancelledError(tool);
    switch (tool) {
      case 'word.get_text':
        data = await getText(args);
        break;
      case 'word.get_paragraph':
        data = await getParagraph(args);
        break;
      case 'word.get_outline':
        data = await getOutline(args);
        break;
      case 'word.find_text':
        data = await findText(args);
        break;
      case 'word.get_selection':
        data = await getSelection(args);
        break;
      case 'word.insert_paragraph':
        data = await insertParagraph(args);
        break;
      case 'word.insert_heading':
        data = await insertHeading(args);
        break;
      case 'word.insert_table':
        data = await insertTable(args);
        break;
      case 'word.insert_image':
        data = await insertImage(args);
        break;
      case 'word.insert_page_break':
        data = await insertPageBreak(args);
        break;
      case 'word.insert_list':
        data = await insertList(args);
        break;
      case 'word.replace_text':
        data = await replaceText(args);
        break;
      case 'word.update_paragraph':
        data = await updateParagraph(args);
        break;
      case 'word.delete_range':
        data = await deleteRange(args);
        break;
      case 'word.apply_formatting':
        data = await applyFormatting(args);
        break;
      case 'word.read_table':
        data = await readTable(args);
        break;
      case 'word.update_cell':
        data = await updateCell(args);
        break;
      case 'word.add_row':
        data = await addRow(args);
        break;
      case 'word.add_column':
        data = await addColumn(args);
        break;
      case 'word.format_cell':
        data = await formatCell(args);
        break;
      case 'word.set_heading_level':
        data = await setHeadingLevel(args);
        break;
      case 'word.apply_style':
        data = await applyStyle(args);
        break;
      case 'word.add_comment':
        data = await addComment(args);
        break;
      case 'word.resolve_comment':
        data = await resolveComment(args);
        break;
      case 'word.accept_change':
        data = await mutateTrackedChange(args, 'accept');
        break;
      case 'word.reject_change':
        data = await mutateTrackedChange(args, 'reject');
        break;
      case 'word._get_comments':
        data = await getComments(args);
        break;
      case 'word._get_tracked_changes':
        data = await getTrackedChanges(args);
        break;
      case 'word._get_structure':
        data = await getStructure(args);
        break;
      case 'word.save':
        data = await saveDocument(args);
        break;
      default:
        throw Object.assign(new Error(`Unsupported tool ${tool}`), { officeMcpCode: 'HOST_CAPABILITY_UNAVAILABLE' });
    }
    if (cancelledRequests.delete(message.id)) throw cancelledError(tool);
    reply(message.id, { ok: true, data, elapsed_ms: Math.round(performance.now() - started) });
  }

  function cancelledError(tool) {
    return Object.assign(new Error(`Tool ${tool} was cancelled.`), { officeMcpCode: 'CANCELLED', partialEffect: 'unknown' });
  }

  async function getText(args) {
    return Word.run(async (context) => {
      const offset = args.offset ?? 0;
      const limit = args.limit ?? 200;
      const includeMetadata = args.include_metadata ?? false;
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text,items/style');
      await context.sync();
      const items = paragraphs.items;
      const selected = items.slice(offset, offset + limit);
      const text = selected.map((paragraph) => paragraph.text).join('\n');
      const data = {
        text,
        paragraph_count: items.length,
        returned_paragraphs: { offset, limit },
        truncated: offset + limit < items.length,
        untrusted_source: true
      };
      if (includeMetadata) {
        data.paragraphs = selected.map((paragraph, index) => ({
          index: offset + index,
          text: paragraph.text,
          style: paragraph.style || null,
          level: headingLevelFromStyle(paragraph.style)
        }));
      }
      return data;
    });
  }

  async function getParagraph(args) {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text,items/style');
      await context.sync();
      const paragraph = paragraphs.items[args.index];
      if (!paragraph) throw Object.assign(new Error(`Paragraph index ${args.index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
      return { index: args.index, text: paragraph.text, style: paragraph.style || null, untrusted_source: true };
    });
  }

  async function getOutline(args) {
    return Word.run(async (context) => {
      const maxLevel = args.max_level ?? 6;
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text,items/style');
      await context.sync();
      const flat = [];
      paragraphs.items.forEach((paragraph, index) => {
        const level = headingLevelFromStyle(paragraph.style);
        if (level > 0 && level <= maxLevel) {
          flat.push({ text: paragraph.text, level, paragraph_index: index, children: [] });
        }
      });
      return { outline: nestOutline(flat), headings: flat.map(({ children, ...heading }) => heading), untrusted_source: true };
    });
  }

  async function findText(args) {
    return Word.run(async (context) => {
      const ranges = context.document.body.search(args.query, {
        matchCase: args.match_case ?? false,
        matchWholeWord: args.whole_word ?? false,
        matchWildcards: args.wildcards ?? false
      });
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      ranges.load('items/text');
      await context.sync();
      const limit = args.limit ?? 50;
      const limitedRanges = ranges.items.slice(0, limit);
      const metadata = mapMatchesToParagraphs(paragraphs.items, args.query, args.match_case ?? false, limit, limitedRanges);
      const matches = metadata.map((match, index) => ({
        paragraph_index: match.paragraph_index,
        occurrence_in_paragraph: match.occurrence_in_paragraph,
        text: limitedRanges[index]?.text ?? match.text,
        snippet: match.snippet
      }));
      return { matches, count: matches.length, truncated: ranges.items.length > limit, untrusted_source: true };
    });
  }

  async function getSelection() {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      selection.load('text');
      paragraphs.load('items');
      await context.sync();
      return {
        text: selection.text,
        paragraph_count: paragraphs.items.length,
        is_empty: selection.text.length === 0,
        untrusted_source: true
      };
    });
  }

  async function insertParagraph(args) {
    return Word.run(async (context) => {
      if (args.anchor.kind === 'start_of_document') {
        const inserted = context.document.body.insertParagraph(args.text, Word.InsertLocation.start);
        if (args.style) inserted.style = args.style;
        if (args.formatting) applyRunFormatting(inserted.font, args.formatting);
        await context.sync();
        return { inserted: true };
      }
      if (args.anchor.kind === 'end_of_document') {
        const inserted = context.document.body.insertParagraph(args.text, Word.InsertLocation.end);
        if (args.style) inserted.style = args.style;
        if (args.formatting) applyRunFormatting(inserted.font, args.formatting);
        await context.sync();
        return { inserted: true };
      }
      const target = await resolveAnchor(context, args.anchor);
      let inserted;
      switch (args.anchor.kind) {
        case 'start_of_document':
        case 'before_paragraph_index':
        case 'before_text':
          inserted = target.insertParagraph(args.text, Word.InsertLocation.before);
          break;
        case 'selection':
        case 'end_of_document':
        case 'paragraph_index':
        case 'after_paragraph_index':
        case 'after_text':
        default:
          inserted = target.insertParagraph(args.text, Word.InsertLocation.after);
          break;
      }
      if (args.style) inserted.style = args.style;
      if (args.formatting) applyRunFormatting(inserted.font, args.formatting);
      await context.sync();
      return { inserted: true };
    });
  }

  async function insertHeading(args) {
    const inserted = await insertParagraph({ text: args.text, anchor: args.anchor, style: `Heading ${args.level}` });
    return { ...inserted, level: args.level };
  }

  async function insertTable(args) {
    validateTableData(args.rows, args.cols, args.data);
    return Word.run(async (context) => {
      const values = args.data ?? blankTable(args.rows, args.cols);
      let table;
      if (args.anchor.kind === 'start_of_document') {
        table = context.document.body.insertTable(args.rows, args.cols, Word.InsertLocation.start, values);
      } else if (args.anchor.kind === 'end_of_document') {
        table = context.document.body.insertTable(args.rows, args.cols, Word.InsertLocation.end, values);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        const location = isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after;
        table = target.insertTable(args.rows, args.cols, location, values);
      }
      if (args.style) table.style = args.style;
      await context.sync();
      return { inserted: true, rows: args.rows, cols: args.cols, header_row: args.header_row ?? false };
    });
  }

  async function insertImage(args) {
    return Word.run(async (context) => {
      const base64 = args.image?.base64;
      if (!base64) throw Object.assign(new Error('word.insert_image requires base64 image data after daemon preprocessing.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      let picture;
      if (args.anchor.kind === 'start_of_document') {
        picture = context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        picture = context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        picture = target.insertInlinePictureFromBase64(base64, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      if (args.alt_text) picture.altTextDescription = args.alt_text;
      if (typeof args.width_pt === 'number') picture.width = args.width_pt;
      if (typeof args.height_pt === 'number') picture.height = args.height_pt;
      await context.sync();
      return { inserted: true, byte_length: args.image.byte_length ?? null, mime_type: args.image.mime_type ?? null };
    });
  }


  async function insertPageBreak(args) {
    return Word.run(async (context) => {
      if (args.anchor.kind === 'start_of_document') {
        context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        target.insertBreak(Word.BreakType.page, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      await context.sync();
      return { inserted: true, break_type: 'page' };
    });
  }

  async function insertList(args) {
    return Word.run(async (context) => {
      const text = args.items.join('\n');
      let range;
      if (args.anchor.kind === 'start_of_document') {
        range = context.document.body.insertText(text, Word.InsertLocation.start);
      } else if (args.anchor.kind === 'end_of_document') {
        range = context.document.body.insertText(text, Word.InsertLocation.end);
      } else {
        const target = await resolveAnchor(context, args.anchor);
        range = target.insertText(text, isBeforeAnchor(args.anchor) ? Word.InsertLocation.before : Word.InsertLocation.after);
      }
      range.paragraphs.load('items');
      await context.sync();
      const paragraphs = range.paragraphs.items;
      const level = args.level ?? 0;
      let list = null;
      for (let index = 0; index < paragraphs.length; index++) {
        const paragraph = paragraphs[index];
        paragraph.style = 'List Paragraph';
        if (index === 0) {
          list = paragraph.startNewList();
          if ((args.kind ?? 'bulleted') === 'numbered') {
            list.setLevelNumbering(level, Word.ListNumbering.arabic);
          } else {
            list.setLevelBullet(level, Word.ListBullet.solid);
          }
        } else if (list) {
          list.load('id');
          await context.sync();
          paragraph.attachToList(list.id, level);
        }
        if (level > 0) paragraph.leftIndent = 18 * level;
      }
      await context.sync();
      return { inserted_items: args.items.length, kind: args.kind ?? 'bulleted', level: args.level ?? 0 };
    });
  }

  async function replaceText(args) {
    return Word.run(async (context) => {
      const scope = args.scope || {};
      const searchRoot = scope.selection_only ? context.document.getSelection() : context.document.body;
      const ranges = searchRoot.search(args.find, {
        matchCase: args.match_case ?? false,
        matchWholeWord: args.whole_word ?? false,
        matchWildcards: args.wildcards ?? false
      });
      ranges.load('items/text');
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();

      const limit = args.limit ?? 500;
      const limitedRanges = ranges.items.slice(0, limit);
      const paragraphMatches = mapMatchesToParagraphs(paragraphs.items, args.find, args.match_case ?? false, limit, limitedRanges);
      const filtered = filterReplaceRanges(limitedRanges, paragraphMatches, scope.paragraph_range);
      filtered.skipped_count += Math.max(0, ranges.items.length - limitedRanges.length);
      if (!args.partial_ok && scope.paragraph_range && filtered.skipped_count > 0) {
        throw Object.assign(new Error('replace_text scope excluded one or more matches. Pass partial_ok to replace only scoped matches.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      }
      const matches = filtered.matches.map((match) => ({
        paragraph_index: match.paragraph_index,
        occurrence_in_paragraph: match.occurrence_in_paragraph,
        text: match.range.text,
        snippet: match.snippet
      }));
      if (args.dry_run) return { replaced_count: 0, matches, dry_run: true, skipped_count: filtered.skipped_count };
      for (const match of filtered.matches) {
        match.range.insertText(args.replace, Word.InsertLocation.replace);
      }
      await context.sync();
      return { replaced_count: filtered.matches.length, matches, skipped_count: filtered.skipped_count };
    });
  }

  function mapMatchesToParagraphs(paragraphs, find, matchCase, limit, ranges) {
    const rangeTexts = ranges ? Array.from(ranges, (range) => range.text || '') : null;
    const query = matchCase ? find : find.toLowerCase();
    const matches = [];
    for (let i = 0; i < paragraphs.length && matches.length < limit; i++) {
      const text = paragraphs[i].text || '';
      const haystack = matchCase ? text : text.toLowerCase();
      let from = 0;
      let occurrence = 0;
      while (matches.length < limit) {
        const currentText = rangeTexts?.[matches.length] || find;
        const currentQuery = matchCase ? currentText : currentText.toLowerCase();
        const index = haystack.indexOf(rangeTexts ? currentQuery : query, from);
        if (index < 0) break;
        occurrence += 1;
        matches.push({
          paragraph_index: i,
          occurrence_in_paragraph: occurrence,
          text: text.slice(index, index + currentText.length),
          snippet: text.slice(Math.max(0, index - 40), Math.min(text.length, index + currentText.length + 40))
        });
        from = index + Math.max(1, currentQuery.length);
        if (rangeTexts && matches.length >= rangeTexts.length) return matches;
      }
    }
    return matches;
  }

  function filterReplaceRanges(ranges, paragraphMatches, paragraphRange) {
    const selected = [];
    let skipped = 0;
    for (let i = 0; i < ranges.length; i++) {
      const metadata = paragraphMatches[i] || { paragraph_index: null, occurrence_in_paragraph: null, snippet: ranges[i].text };
      const inRange = !paragraphRange || (metadata.paragraph_index >= paragraphRange[0] && metadata.paragraph_index <= paragraphRange[1]);
      if (inRange) selected.push({ range: ranges[i], ...metadata });
      else skipped += 1;
    }
    return { matches: selected, skipped_count: skipped };
  }

  async function updateParagraph(args) {
    return Word.run(async (context) => {
      const paragraph = await getParagraphByIndex(context, args.index);
      paragraph.insertText(args.text, Word.InsertLocation.replace);
      await context.sync();
      return { updated_paragraph_index: args.index };
    });
  }

  async function deleteRange(args) {
    return Word.run(async (context) => {
      const target = args.extent === 'selection' ? context.document.getSelection() : await resolveAnchor(context, args.anchor);
      if (args.extent === 'sentence') {
        const sentence = await rangeSentenceOrSelf(context, target);
        sentence.delete();
        await context.sync();
        return { deleted: true, extent: 'sentence' };
      }
      if (target.paragraphs && (args.extent ?? 'paragraph') === 'paragraph') {
        const paragraphs = target.paragraphs;
        paragraphs.load('items');
        await context.sync();
        if (paragraphs.items[0]) paragraphs.items[0].delete();
      } else {
        target.delete();
      }
      await context.sync();
      return { deleted: true, extent: args.extent ?? 'paragraph' };
    });
  }

  async function applyFormatting(args) {
    return Word.run(async (context) => {
      const target = args.extent === 'selection' ? context.document.getSelection() : await resolveAnchor(context, args.anchor);
      const range = args.extent === 'sentence' ? await rangeSentenceOrSelf(context, target) : target;
      applyRunFormatting(range.font, args.formatting);
      await context.sync();
      return { formatted: true };
    });
  }

  async function rangeSentenceOrSelf(context, target) {
    target.load('text');
    await context.sync();
    const sourceText = target.text || '';
    const paragraphs = target.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const paragraph = paragraphs.items[0];
    if (!paragraph) return target;
    const sentences = paragraph.split(['.', '!', '?'], false, true);
    sentences.load('items/text');
    await context.sync();
    if (!sourceText) return sentences.items[0] || target;
    return sentences.items.find((sentence) => sentence.text.includes(sourceText)) || sentences.items[0] || target;
  }

  async function readTable(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.load('rowCount,columnCount,values');
      await context.sync();
      const data = table.values ?? [];
      return { rows: table.rowCount ?? data.length, cols: table.columnCount ?? (data[0]?.length ?? 0), data, header_row: false, untrusted_source: true };
    });
  }

  async function updateCell(args) {
    return Word.run(async (context) => {
      const cell = await getTableCell(context, args.table_index, args.row, args.col);
      cell.value = args.text;
      if (args.formatting) applyRunFormatting(cell.body.getRange().font, args.formatting);
      await context.sync();
      return { updated: true, table_index: args.table_index, row: args.row, col: args.col };
    });
  }

  async function addRow(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.load('rowCount,columnCount,values');
      await context.sync();
      const rowCount = table.rowCount ?? table.values.length;
      const columnCount = table.columnCount ?? (table.values[0]?.length ?? 0);
      const values = args.values ?? blankRow(columnCount);
      if (values.length !== columnCount) throw Object.assign(new Error('Row values length must match table column count.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      const rowIndex = args.index ?? rowCount;
      if (rowIndex >= rowCount) {
        table.getCell(rowCount - 1, 0).insertRows(Word.InsertLocation.after, 1, [values]);
      } else {
        table.getCell(rowIndex, 0).insertRows(Word.InsertLocation.before, 1, [values]);
      }
      await context.sync();
      return { added_row_index: rowIndex };
    });
  }

  async function addColumn(args) {
    return Word.run(async (context) => {
      const table = await getTableByIndex(context, args.table_index);
      table.load('rowCount,columnCount,values');
      await context.sync();
      const rowCount = table.rowCount ?? table.values.length;
      const columnCount = table.columnCount ?? (table.values[0]?.length ?? 0);
      const values = args.values ?? blankColumn(rowCount);
      if (values.length !== rowCount) throw Object.assign(new Error('Column values length must match table row count.'), { officeMcpCode: 'INVALID_ARGUMENT' });
      const colIndex = args.index ?? columnCount;
      const matrix = values.map((value) => [value]);
      if (colIndex >= columnCount) {
        table.getCell(0, columnCount - 1).insertColumns(Word.InsertLocation.after, 1, matrix);
      } else {
        table.getCell(0, colIndex).insertColumns(Word.InsertLocation.before, 1, matrix);
      }
      await context.sync();
      return { added_column_index: colIndex };
    });
  }

  async function formatCell(args) {
    return Word.run(async (context) => {
      const cell = await getTableCell(context, args.table_index, args.row, args.col);
      if (args.background_color) cell.shadingColor = args.background_color;
      if (args.horizontal_alignment) cell.horizontalAlignment = mapHorizontalAlignment(args.horizontal_alignment);
      if (args.vertical_alignment) cell.verticalAlignment = mapVerticalAlignment(args.vertical_alignment);
      if (typeof args.padding_pt === 'number') {
        cell.setCellPadding(Word.CellPaddingLocation.top, args.padding_pt);
        cell.setCellPadding(Word.CellPaddingLocation.bottom, args.padding_pt);
        cell.setCellPadding(Word.CellPaddingLocation.left, args.padding_pt);
        cell.setCellPadding(Word.CellPaddingLocation.right, args.padding_pt);
      }
      if (args.formatting) applyRunFormatting(cell.body.getRange().font, args.formatting);
      await context.sync();
      return { formatted: true, table_index: args.table_index, row: args.row, col: args.col };
    });
  }

  async function setHeadingLevel(args) {
    return Word.run(async (context) => {
      const paragraph = await getParagraphByIndex(context, args.index);
      paragraph.style = args.level === 0 ? 'Normal' : `Heading ${args.level}`;
      await context.sync();
      return { paragraph_index: args.index, level: args.level };
    });
  }

  async function applyStyle(args) {
    return Word.run(async (context) => {
      const target = await resolveAnchor(context, args.anchor);
      target.style = args.style;
      await context.sync();
      return { styled: true, style: args.style };
    });
  }

  async function getComments() {
    return Word.run(async (context) => {
      const comments = context.document.comments;
      comments.load('items/id,items/content,items/resolved,items/authorName,items/creationDate');
      await context.sync();
      return {
        comments: comments.items.map((comment, index) => ({
          index,
          comment_id: comment.id,
          content: comment.content,
          resolved: comment.resolved,
          author: comment.authorName || null,
          created_at: dateToIso(comment.creationDate),
          untrusted_source: true
        })),
        count: comments.items.length,
        untrusted_source: true
      };
    });
  }

  async function getTrackedChanges() {
    return Word.run(async (context) => {
      const changes = context.document.body.getTrackedChanges();
      changes.load('items/author,items/date,items/type,items/text');
      await context.sync();
      return {
        changes: changes.items.map((change, index) => ({
          index,
          author: change.author || null,
          date: dateToIso(change.date),
          type: change.type || null,
          text: change.text || '',
          fingerprint: trackedChangeFingerprint(change, index),
          untrusted_source: true
        })),
        count: changes.items.length,
        untrusted_source: true
      };
    });
  }

  async function getStructure() {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      const tables = context.document.body.tables;
      paragraphs.load('items/text,items/style,items/isListItem');
      tables.load('items/rowCount,items/columnCount');
      await context.sync();
      const listMetadata = new Map();
      const listObjects = [];
      paragraphs.items.forEach((paragraph, index) => {
        if (!paragraph.isListItem) return;
        const listItem = paragraph.listItemOrNullObject;
        const list = paragraph.listOrNullObject;
        listItem.load('isNullObject,level');
        list.load('isNullObject,levelTypes');
        listObjects.push({ index, listItem, list });
      });
      if (listObjects.length > 0) await context.sync();
      for (const item of listObjects) {
        const level = item.listItem.isNullObject ? null : item.listItem.level;
        const levelTypes = item.list.isNullObject ? [] : item.list.levelTypes ?? [];
        listMetadata.set(item.index, { level, kind: listKindFromLevelType(levelTypes[level ?? 0]) });
      }
      const headings = [];
      const lists = [];
      paragraphs.items.forEach((paragraph, index) => {
        const level = headingLevelFromStyle(paragraph.style);
        if (level > 0) headings.push({ text: paragraph.text, level, paragraph_index: index });
        if (paragraph.isListItem) {
          const metadata = listMetadata.get(index) || { level: null, kind: null };
          lists.push({ paragraph_index: index, text: paragraph.text, level: metadata.level, kind: metadata.kind });
        }
      });
      return {
        outline: nestOutline(headings.map((heading) => ({ ...heading, children: [] }))),
        headings,
        lists,
        tables: tables.items.map((table, index) => ({ index, rows: table.rowCount ?? null, cols: table.columnCount ?? null })),
        untrusted_source: true
      };
    });
  }

  async function addComment(args) {
    return Word.run(async (context) => {
      const target = await resolveAnchor(context, args.anchor);
      const comment = target.insertComment(args.text);
      comment.load('id,content,resolved,authorName,creationDate');
      await context.sync();
      return {
        comment_id: comment.id,
        content: comment.content,
        resolved: comment.resolved,
        author: comment.authorName || null,
        created_at: dateToIso(comment.creationDate)
      };
    });
  }

  async function resolveComment(args) {
    return Word.run(async (context) => {
      const comments = context.document.comments;
      comments.load('items/id,items/resolved');
      await context.sync();
      const comment = comments.items.find((item) => item.id === args.comment_id);
      if (!comment) throw Object.assign(new Error(`Comment ${args.comment_id} was not found.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
      comment.resolved = true;
      await context.sync();
      return { comment_id: args.comment_id, resolved: true };
    });
  }

  async function mutateTrackedChange(args, action) {
    return Word.run(async (context) => {
      const changes = context.document.body.getTrackedChanges();
      changes.load('items/author,items/date,items/type,items/text');
      await context.sync();
      const change = changes.items[args.change_index];
      if (!change) throw Object.assign(new Error(`Tracked change index ${args.change_index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
      const fingerprint = trackedChangeFingerprint(change, args.change_index);
      if (fingerprint !== args.expected_fingerprint) {
        throw Object.assign(new Error('Tracked change fingerprint mismatch; re-read track_changes before mutating.'), { officeMcpCode: 'STALE_INDEX' });
      }
      if (action === 'accept') change.accept();
      else change.reject();
      await context.sync();
      return { change_index: args.change_index, fingerprint, action };
    });
  }

  async function saveDocument() {
    return Word.run(async (context) => {
      const document = context.document;
      document.load('saved');
      await context.sync();
      const wasDirty = typeof document.saved === 'boolean' ? !document.saved : null;
      document.save(Word.SaveBehavior.save);
      document.load('saved');
      await context.sync();
      return { saved: true, was_dirty: wasDirty, is_dirty: typeof document.saved === 'boolean' ? !document.saved : null };
    });
  }

  async function resolveAnchor(context, anchor) {
    switch (anchor.kind) {
      case 'selection':
        return context.document.getSelection();
      case 'start_of_document':
      case 'end_of_document':
        return context.document.body;
      case 'paragraph_index':
      case 'before_paragraph_index':
      case 'after_paragraph_index': {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        const paragraph = paragraphs.items[anchor.index];
        if (!paragraph) throw Object.assign(new Error(`Paragraph index ${anchor.index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
        return paragraph;
      }
      case 'after_text':
      case 'before_text': {
        const ranges = context.document.body.search(anchor.text, { matchCase: false });
        ranges.load('items');
        await context.sync();
        const range = ranges.items[(anchor.occurrence ?? 1) - 1];
        if (!range) throw Object.assign(new Error(`Anchor text not found: ${anchor.text}`), { officeMcpCode: 'ANCHOR_NOT_FOUND' });
        return range;
      }
      case 'heading': {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items/text,items/style');
        await context.sync();
        const expectedLevel = anchor.level || null;
        const paragraph = paragraphs.items.find((item) => item.text === anchor.text && headingLevelFromStyle(item.style) > 0 && (!expectedLevel || headingLevelFromStyle(item.style) === expectedLevel));
        if (!paragraph) throw Object.assign(new Error('Heading anchor not found: ' + anchor.text), { officeMcpCode: 'ANCHOR_NOT_FOUND' });
        return paragraph;
      }
      case 'bookmark': {
        const range = context.document.getBookmarkRangeOrNullObject(anchor.name);
        range.load('isNullObject,text');
        await context.sync();
        if (range.isNullObject) throw Object.assign(new Error('Bookmark anchor not found: ' + anchor.name), { officeMcpCode: 'ANCHOR_NOT_FOUND' });
        return range;
      }
      default:
        throw Object.assign(new Error(`Unsupported anchor ${anchor.kind}`), { officeMcpCode: 'INVALID_ARGUMENT' });
    }
  }

  async function getParagraphByIndex(context, index) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load('items');
    await context.sync();
    const paragraph = paragraphs.items[index];
    if (!paragraph) throw Object.assign(new Error(`Paragraph index ${index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
    return paragraph;
  }

  async function getTableByIndex(context, index) {
    const tables = context.document.body.tables;
    tables.load('items');
    await context.sync();
    const table = tables.items[index];
    if (!table) throw Object.assign(new Error(`Table index ${index} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
    return table;
  }

  async function getTableCell(context, tableIndex, row, col) {
    const table = await getTableByIndex(context, tableIndex);
    table.load('rowCount,columnCount');
    await context.sync();
    if (row >= table.rowCount || col >= table.columnCount) {
      throw Object.assign(new Error(`Cell ${row},${col} is out of range.`), { officeMcpCode: 'INDEX_OUT_OF_RANGE' });
    }
    return table.getCell(row, col);
  }

  function isBeforeAnchor(anchor) {
    return anchor.kind === 'before_paragraph_index' || anchor.kind === 'before_text' || anchor.kind === 'start_of_document';
  }

  function applyRunFormatting(font, formatting) {
    if (typeof formatting.bold === 'boolean') font.bold = formatting.bold;
    if (typeof formatting.italic === 'boolean') font.italic = formatting.italic;
    if (typeof formatting.underline === 'boolean') font.underline = formatting.underline ? Word.UnderlineType.single : Word.UnderlineType.none;
    if (typeof formatting.strikethrough === 'boolean') font.strikeThrough = formatting.strikethrough;
    if (formatting.font_name) font.name = formatting.font_name;
    if (typeof formatting.font_size_pt === 'number') font.size = formatting.font_size_pt;
    if (formatting.color) font.color = formatting.color;
    if (formatting.highlight) font.highlightColor = formatting.highlight;
  }

  function headingLevelFromStyle(style) {
    const match = String(style || '').match(/^Heading\s+([1-9])$/i);
    return match ? Number(match[1]) : 0;
  }

  function listKindFromLevelType(levelType) {
    const value = String(levelType || '').toLowerCase();
    if (value === 'number') return 'numbered';
    if (value === 'bullet') return 'bulleted';
    if (value === 'picture') return 'picture';
    return null;
  }

  function nestOutline(flat) {
    const roots = [];
    const stack = [];
    for (const node of flat) {
      const copy = { ...node, children: [] };
      while (stack.length && stack[stack.length - 1].level >= copy.level) stack.pop();
      if (stack.length) stack[stack.length - 1].children.push(copy);
      else roots.push(copy);
      stack.push(copy);
    }
    return roots;
  }

  function validateTableData(rows, cols, data) {
    if (!data) return;
    if (data.length !== rows || data.some((row) => row.length !== cols)) {
      throw Object.assign(new Error('Table data dimensions must match rows and cols.'), { officeMcpCode: 'INVALID_ARGUMENT' });
    }
  }

  function blankTable(rows, cols) {
    return Array.from({ length: rows }, () => blankRow(cols));
  }

  function blankRow(cols) {
    return Array.from({ length: cols }, () => '');
  }

  function blankColumn(rows) {
    return Array.from({ length: rows }, () => '');
  }

  function mapHorizontalAlignment(value) {
    return value === 'center' ? Word.Alignment.centered : value === 'right' ? Word.Alignment.right : Word.Alignment.left;
  }

  function mapVerticalAlignment(value) {
    return value === 'center' ? Word.VerticalAlignment.center : value === 'bottom' ? Word.VerticalAlignment.bottom : Word.VerticalAlignment.top;
  }

  function trackedChangeFingerprint(change, index) {
    const date = change.date instanceof Date ? change.date.toISOString() : String(change.date || '');
    return stableHash([index, change.author || '', date, change.type || '', change.text || ''].join('\u001f'));
  }

  function stableHash(input) {
    let hash = 2166136261;
    for (let i = 0; i < input.length; i += 1) {
      hash ^= input.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, '0');
  }

  function dateToIso(value) {
    if (!value) return null;
    if (value instanceof Date) return value.toISOString();
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? String(value) : parsed.toISOString();
  }

  async function getDocumentInfo() {
    return Word.run(async (context) => {
      const document = context.document;
      document.load('saved');
      await context.sync();
      return {
        title: Office.context.document.url ? fileName(Office.context.document.url) : 'Word document',
        url: Office.context.document.url || null,
        filename: Office.context.document.url ? fileName(Office.context.document.url) : null,
        is_dirty: typeof document.saved === 'boolean' ? !document.saved : null,
        is_read_only: null,
        is_protected: null,
        protection: { kind: null, rights: null, rights_source: 'unavailable' },
        opened_at: new Date().toISOString()
      };
    });
  }

  function probeRequirementSets() {
    const requirements = Office.context.requirements;
    return {
      WordApi: requirements.isSetSupported('WordApi', '1.3') ? '1.3' : null,
      WordApi_1_4: requirements.isSetSupported('WordApi', '1.4') ? '1.4' : null,
      WordApi_1_6: requirements.isSetSupported('WordApi', '1.6') ? '1.6' : null
    };
  }

  function mapError(error, tool) {
    const code = error.officeMcpCode || classifyOfficeError(error);
    return {
      office_mcp_code: code,
      message: error.message || String(error),
      session_id: sessionId,
      tool,
      retriable: Boolean(error.retriable) || code === 'HOST_BUSY' || code === 'TIMEOUT',
      partial_effect: error.partialEffect || 'unknown'
    };
  }

  function classifyOfficeError(error) {
    const code = String(error.code || error.name || '');
    const message = String(error.message || '');
    if (/permission|denied|IRM|rights/i.test(code + message)) return 'IRM_DENIED';
    if (/read.?only/i.test(code + message)) return 'DOCUMENT_READ_ONLY';
    if (/STALE_INDEX/i.test(code + message)) return 'STALE_INDEX';
    return 'GENERIC_FAILURE';
  }

  function reply(id, result) {
    send({ jsonrpc: '2.0', id, result });
  }

  function send(message) {
    if (socket && socket.readyState === WebSocket.OPEN) socket.send(JSON.stringify(message));
  }

  function setStatus(text) {
    statusEl.textContent = text;
  }

  function fileName(path) {
    return path.split(/[\\/]/).pop() || path;
  }
})();




