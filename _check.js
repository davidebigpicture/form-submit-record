const fs = require('fs');
const html = fs.readFileSync('form_submit_tool.html', 'utf8');

// Pull every inline <script> block and run them as one program.
let code = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');

// Test harness appended in the same scope: for every form type, render the grid
// and the detail panel for every record so all code paths actually execute.
code += `
    ;(function () {
        const out = [];
        const grids = [];
        Object.keys(recordSets).forEach((type) => {
            activeType = type;
            activeId = null;
            renderGrid();
            const gbody = grid.innerHTML;
            grids.push({ type, rows: (gbody.match(/<tr /g) || []).length, count: recordSets[type].length, hasUndefined: /undefined/.test(gbody) });
            recordSets[type].forEach((r) => {
                activeId = r.id;
                renderDetail();
                const html = detailPanel.innerHTML;
                out.push({ type, id: r.id, len: html.length, opCards: (html.match(/op-card/g) || []).length, hasUndefined: /undefined/.test(html) });
            });
        });
        activeType = "RENEW";
        activeId = null;
        renderGrid();
        globalThis.__detailReport = out;
        globalThis.__gridReport = grids;
    })();
`;

const stubs = {};
const makeStub = () => ({
    innerHTML: '',
    textContent: '',
    hidden: false,
    dataset: {},
    style: {},
    classList: { add() {}, remove() {}, toggle() {} },
    addEventListener() {},
    removeEventListener() {},
    setAttribute() {},
    getAttribute() { return null; },
    appendChild() {},
    contains() { return false; },
    scrollIntoView() {},
    querySelector() { return makeStub(); },
    querySelectorAll() { return []; },
});

global.HTMLInputElement = function () {};
global.HTMLAnchorElement = function () {};
global.document = {
    getElementById: (id) => (stubs[id] = stubs[id] || makeStub()),
    createElement: () => makeStub(),
    addEventListener() {},
    head: makeStub(),
    querySelector: () => makeStub(),
    querySelectorAll: () => [],
};

try {
    new Function(code)();
    const body = stubs.gridBody ? stubs.gridBody.innerHTML : '';
    const rows = (body.match(/<tr /g) || []).length;
    const chips = (body.match(/class="chip/g) || []).length;
    console.log('grid HTML length:', body.length);
    console.log('row count:', rows);
    console.log('chip count:', chips);
    console.log('first 600 chars:\n', body.slice(0, 600));
    if (rows === 0) {
        console.log('FAIL: no rows rendered');
        process.exit(1);
    }
    console.log('\nOK: grid rendered', rows, 'rows');

    const grids = globalThis.__gridReport || [];
    console.log('\nGrids per form type:');
    let gridBad = false;
    grids.forEach((g) => {
        const flag = g.hasUndefined ? ' <-- contains "undefined"' : '';
        if (g.hasUndefined || g.rows !== g.count || g.rows === 0) gridBad = true;
        console.log(`  ${g.type}: ${g.rows} rows (expected ${g.count})${flag}`);
    });
    if (gridBad) {
        console.log('FAIL: a form-type grid rendered the wrong number of rows or leaked "undefined"');
        process.exit(1);
    }

    const report = globalThis.__detailReport || [];
    console.log('\nDetail panels rendered:', report.length);
    let bad = false;
    report.forEach((d) => {
        const flag = d.hasUndefined ? ' <-- contains "undefined"' : '';
        if (d.hasUndefined || d.opCards === 0) bad = true;
        console.log(`  ${d.type} #${d.id}: ${d.len} chars, ${d.opCards} op-cards${flag}`);
    });
    if (bad) {
        console.log('FAIL: a detail panel had no op-cards or leaked "undefined"');
        process.exit(1);
    }
    console.log('\nOK: all detail panels rendered with op-cards');
} catch (e) {
    console.log('RUNTIME ERROR:', e.message);
    console.log(e.stack.split('\n').slice(0, 8).join('\n'));
    process.exit(1);
}
