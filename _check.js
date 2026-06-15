const fs = require('fs');
const html = fs.readFileSync('form_submit_tool.html', 'utf8');
const m = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const stubs = {};
const makeStub = () => ({ innerHTML: '', addEventListener: () => {}, querySelectorAll: () => [], scrollIntoView: () => {} });
global.document = { getElementById: (id) => (stubs[id] = stubs[id] || makeStub()) };
global.HTMLInputElement = function () {};
global.HTMLAnchorElement = function () {};
try {
    new Function(m)();
    console.log('grid HTML length:', stubs.gridBody.innerHTML.length);
    console.log('row count:', (stubs.gridBody.innerHTML.match(/<tr /g) || []).length);
    console.log('first 800 chars:\n', stubs.gridBody.innerHTML.slice(0, 800));
} catch (e) {
    console.log('RUNTIME ERROR:', e.message);
    console.log(e.stack.split('\n').slice(0, 8).join('\n'));
}
