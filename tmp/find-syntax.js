const fs = require('fs');
const html = fs.readFileSync('form_submit_tool.html', 'utf8');
const code = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]).join('\n');
const lines = code.split('\n');
for (let i = 0; i < lines.length; i++) {
    try {
        new Function(lines.slice(0, i + 1).join('\n'));
    } catch (e) {
        if (String(e.message).includes('Unexpected token')) {
            console.log('First error near line', i + 1);
            for (let j = Math.max(0, i - 3); j <= Math.min(lines.length - 1, i + 3); j++) {
                console.log(String(j + 1).padStart(5), lines[j]);
            }
            break;
        }
    }
}
