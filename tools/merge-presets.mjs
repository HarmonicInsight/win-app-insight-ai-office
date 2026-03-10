// Node.js script to extract presets from sister apps and merge into IAOF
import { readFileSync, writeFileSync } from 'fs';

const slideFile = 'C:/dev/win-app-insight-slide/src/InsightOfficeSlide/Services/PromptPresetService.cs';
const sheetFile = 'C:/dev/win-app-insight-sheet/src/HarmonicSheet.App/Helpers/PromptPresetService.cs';
const docFile   = 'C:/dev/win-app-insight-doc/src/HarmonicDoc.Core/Services/PromptPresetService.cs';

function extractPresetBlock(filePath) {
    const content = readFileSync(filePath, 'utf-8');
    // Find "private static List<UserPromptPreset> GetBuiltInPresets()" and extract the return block
    const marker = 'private static List<UserPromptPreset> GetBuiltInPresets()';
    const idx = content.indexOf(marker);
    if (idx === -1) { console.error(`ERROR: marker not found in ${filePath}`); return ''; }

    // Find "return" after the marker, then find the opening "["
    const returnIdx = content.indexOf('return', idx);
    const openBracket = content.indexOf('[', returnIdx);
    if (openBracket === -1) return '';

    // Find matching close bracket
    let depth = 1;
    let i = openBracket + 1;
    while (i < content.length && depth > 0) {
        if (content[i] === '[') depth++;
        else if (content[i] === ']') depth--;
        i++;
    }
    return content.substring(openBracket + 1, i - 1).trim();
}

function transformPresets(block, idPrefix, categoryPrefix, emoji) {
    let result = block;

    // Replace builtin_ IDs with prefix
    result = result.replace(/Id = "builtin_/g, `Id = "${idPrefix}`);

    // Replace Category = "xxx" with Category = "emoji prefix: xxx"
    result = result.replace(/Category\s*=\s*"([^"]+)"/g, (_, cat) =>
        `Category = "${emoji} ${categoryPrefix}: ${cat}"`);

    // Remove IsDefault = true
    result = result.replace(/,?\s*IsDefault\s*=\s*true\s*,?/g, (match) => {
        if (match.startsWith(',') && match.endsWith(',')) return ',';
        if (match.startsWith(',')) return '';
        if (match.endsWith(',')) return '';
        return '';
    });

    return result;
}

console.log('Reading source files...');

const slideBlock = extractPresetBlock(slideFile);
const sheetBlock = extractPresetBlock(sheetFile);
const docBlock   = extractPresetBlock(docFile);

console.log(`SLIDE block: ${slideBlock.length} chars`);
console.log(`SHEET block: ${sheetBlock.length} chars`);
console.log(`DOC block: ${docBlock.length} chars`);

const slidePresets = transformPresets(slideBlock, 'slide_', 'Slide', '📊');
const sheetPresets = transformPresets(sheetBlock, 'sheet_', 'Sheet', '📗');
const docPresets   = transformPresets(docBlock, 'doc_', 'Doc', '📄');

// Count presets
const count = (s) => (s.match(/\bnew\(\)/g) || []).length;
console.log(`SLIDE presets: ${count(slidePresets)}`);
console.log(`SHEET presets: ${count(sheetPresets)}`);
console.log(`DOC presets: ${count(docPresets)}`);
console.log(`TOTAL imported: ${count(slidePresets) + count(sheetPresets) + count(docPresets)}`);

// Write merged fragment
const output = `        // ════════════════════════════════════════════════════════════
        // 📊 InsightSlide からインポート
        // ════════════════════════════════════════════════════════════
${slidePresets}

        // ════════════════════════════════════════════════════════════
        // 📗 InsightSheet からインポート
        // ════════════════════════════════════════════════════════════
${sheetPresets}

        // ════════════════════════════════════════════════════════════
        // 📄 InsightDoc からインポート
        // ════════════════════════════════════════════════════════════
${docPresets}
`;

writeFileSync('C:/dev/win-app-insight-ai-office/tools/merged-presets-fragment.cs', output, 'utf-8');
console.log(`\nFragment written. Total chars: ${output.length}`);
