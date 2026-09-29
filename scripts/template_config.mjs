// Read the same template identity as the Python application.
import fs from 'node:fs/promises';
import { fileURLToPath } from 'node:url';

const directory = new URL('../assets/templates/', import.meta.url);
const metadata = JSON.parse(await fs.readFile(new URL('manifest.json', directory), 'utf8'));
export const templateVersion = metadata.version;
export const templatePath = fileURLToPath(new URL(metadata.filename, directory));
