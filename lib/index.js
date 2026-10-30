/**
 * Bundled group-meeting-slides skill provider for DeepSeek Harness.
 *
 * The package ships the skill under "group-meeting-slides/" and registers one
 * skill provider on "ctx.skills". Discovery reads only the frontmatter of
 * "SKILL.md"; the instruction body and the "references/*.md" files are read on
 * demand, and the candidate carries the skill directory as its resource base so
 * relative references resolve. The Python helpers in "scripts/" are executed by
 * the agent against that same directory.
 *
 * @module dsh-group-meeting-slides
 */

import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

/** Provider name registered on "ctx.skills" (the runtime name "runtime" is reserved). */
const PROVIDER_NAME = "group-meeting-slides";

/** Source label carried by every candidate and definition. */
const SOURCE = "group-meeting-slides";

/** Standard rank for packaged skill providers and bundled roots (BUNDLED_SKILL_RANK). */
const RANK = 600;

/** Absolute asset root, resolved from this module inside the installed package. */
const ASSET_ROOT = fileURLToPath(new URL("../group-meeting-slides/", import.meta.url));

/** Stable Loader identity. */
const name = "group-meeting-slides";

/** The skill registry must exist before "apply" runs. */
const inject = ["skills"];

/** The invocation policy: both the model and the user may load this skill. */
const INVOCATION = { modelInvocable: true, userInvocable: true };

/** The YAML frontmatter block at the start of a skill file. */
const FRONTMATTER = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/u;

/** One "key: value" scalar line inside the frontmatter block. */
const FIELD = /^([A-Za-z][A-Za-z0-9_-]*):[ \t]*(.*)$/u;

/** A complete single-line double-quoted YAML scalar. */
const DOUBLE_QUOTED = /^"((?:[^"\\]|\\.)*)"$/u;

/** A complete single-line single-quoted YAML scalar. */
const SINGLE_QUOTED = /^'((?:[^']|'')*)'$/u;

/** Absolute path of the packaged SKILL.md. */
const SKILL_PATH = ASSET_ROOT + "SKILL.md";

/**
 * Read one YAML scalar value.
 *
 * Only single-line scalars are supported, which is all this package ships. A
 * value that opens but does not close on the same line is rejected instead of
 * being silently truncated.
 *
 * @param raw - the trimmed value text after the colon.
 * @param where - a label used in error messages.
 * @returns the decoded scalar.
 */
function parseScalar(raw, where) {
	if (raw.length === 0) return "";
	const double = DOUBLE_QUOTED.exec(raw);
	if (double !== null) return double[1].replace(/\\([\s\S])/gu, (_, c) => (c === "n" ? "\n" : c === "t" ? "\t" : c));
	const single = SINGLE_QUOTED.exec(raw);
	if (single !== null) return single[1].replace(/''/gu, "'");
	if (raw.startsWith("\"") || raw.startsWith("'")) throw new Error(where + ": unterminated quoted scalar");
	return raw;
}

/**
 * Parse one SKILL.md file into the registry fields and the instruction body.
 *
 * @param raw - the complete file contents.
 * @param path - the file path, used in error messages.
 * @returns the validated name and description plus the body without frontmatter.
 * @throws when the frontmatter, the name, or the description is missing or malformed.
 */
export function parseSkillFile(raw, path) {
	if (typeof raw !== "string") throw new TypeError(path + ": SKILL.md content must be a string");
	const text = raw.charCodeAt(0) === 0xfeff ? raw.slice(1) : raw;
	const match = FRONTMATTER.exec(text);
	if (match === null) throw new Error(path + ": missing YAML frontmatter");
	const fields = new Map();
	for (const line of match[1].split(/\r?\n/u)) {
		const entry = FIELD.exec(line);
		if (entry === null) continue;
		fields.set(entry[1], parseScalar(entry[2].trim(), path + ': "' + entry[1] + '"'));
	}
	const skillName = fields.get("name");
	const description = fields.get("description");
	if (typeof skillName !== "string" || !/^[a-z0-9]+(?:-[a-z0-9]+)*$/u.test(skillName)) {
		throw new Error(path + ': "name" must be kebab-case, got ' + JSON.stringify(skillName));
	}
	if (typeof description !== "string" || description.length === 0) {
		throw new Error(path + ': "description" is required');
	}
	const body = text.slice(match[0].length).trim().replace(/\r\n?/gu, "\n");
	return { name: skillName, description, content: body };
}

/**
 * Read and validate the packaged skill, annotating failures with the plugin
 * name so a broken installation is diagnosable from the host log.
 *
 * @param signal - optional abort signal from the registry.
 * @returns the parsed skill file.
 */
async function loadSkill(signal) {
	let raw;
	try {
		raw = await readFile(SKILL_PATH, { encoding: "utf8", signal });
	} catch (cause) {
		throw new Error("group-meeting-slides: cannot read " + SKILL_PATH + ": " + cause.message, { cause });
	}
	try {
		return parseSkillFile(raw, SKILL_PATH);
	} catch (cause) {
		throw new Error("group-meeting-slides: " + cause.message, { cause });
	}
}

/**
 * Register the bundled skill provider on "ctx.skills".
 *
 * The returned disposer of "registerProvider" is owned by the calling fiber,
 * so plugin unload removes the provider without an extra "ctx.effect" wrapper.
 *
 * @param ctx - Cordis context carrying the skill registry.
 */
function apply(ctx) {
	const provider = {
		name: PROVIDER_NAME,
		async list(options) {
			const skill = await loadSkill(options?.signal);
			return [{
				name: skill.name,
				description: skill.description,
				invocation: { ...INVOCATION },
				source: SOURCE,
				provider: PROVIDER_NAME,
				rank: RANK,
				locator: SKILL_PATH,
				path: SKILL_PATH,
				resourceBase: { kind: "directory", path: ASSET_ROOT }
			}];
		},
		async get(candidate, options) {
			if (candidate.name !== name) {
				throw new Error('group-meeting-slides: unknown skill "' + candidate.name + '"');
			}
			const skill = await loadSkill(options?.signal);
			return {
				name: skill.name,
				description: skill.description,
				invocation: { ...INVOCATION },
				source: SOURCE,
				provider: PROVIDER_NAME,
				resourceBase: { kind: "directory", path: ASSET_ROOT },
				path: SKILL_PATH,
				content: skill.content
			};
		}
	};
	ctx.skills.registerProvider(() => provider);
}

export { apply, inject, name };