const { existsSync, readFileSync } = require("node:fs");
const path = require("node:path");

const ALLOWLIST = new Set(["MIT", "MIT-0", "ISC", "APACHE-2.0", "ZLIB", "BSD-2-CLAUSE", "BSD-3-CLAUSE", "0BSD"]);
const DISPLAY_NAMES = { jszip: "JSZip" };

function loadLicenceRecords(root) {
  const filePath = path.join(root, "scripts", "licence-elections.json");
  if (!existsSync(filePath)) return { elections: {}, approvals: {} };
  const parsed = JSON.parse(readFileSync(filePath, "utf8"));
  if (parsed && typeof parsed === "object" && parsed.elections && typeof parsed.elections === "object") {
    return {
      elections: parsed.elections,
      approvals: parsed.approvals && typeof parsed.approvals === "object" ? parsed.approvals : {},
    };
  }
  return { elections: parsed && typeof parsed === "object" ? parsed : {}, approvals: {} };
}

function splitIdentity(identity) {
  const separator = identity.lastIndexOf("@");
  if (separator <= 0 || separator === identity.length - 1) {
    throw new Error(`Licence identity is malformed: ${identity}`);
  }
  return { name: identity.slice(0, separator), version: identity.slice(separator + 1) };
}

function displayName(packageName) {
  return DISPLAY_NAMES[packageName] || packageName;
}

function tokenise(expression) {
  const source = expression.trim();
  const tokens = [];
  const regex = /\s*(\(|\)|AND|OR|WITH|[A-Za-z0-9.+-]+)\s*/gi;
  let cursor = 0;
  let match;
  while ((match = regex.exec(source))) {
    if (match.index !== cursor) {
      throw new Error(`Unrecognised licence expression syntax: ${expression}`);
    }
    const raw = match[1];
    const upper = raw.toUpperCase();
    if (upper === "AND" || upper === "OR" || upper === "WITH" || raw === "(" || raw === ")") {
      tokens.push(upper === "AND" || upper === "OR" || upper === "WITH" ? upper : raw);
    } else {
      tokens.push(raw);
    }
    cursor = regex.lastIndex;
  }
  if (cursor !== source.length || tokens.length === 0) {
    throw new Error(`Unrecognised licence expression syntax: ${expression}`);
  }
  return tokens;
}

function parseExpression(expression) {
  const tokens = tokenise(expression);
  if (tokens.includes("WITH")) {
    throw new Error(`Licence expression uses a WITH exception and needs recorded approval: ${expression}`);
  }
  let index = 0;

  function peek() {
    return tokens[index];
  }

  function take() {
    return tokens[index++];
  }

  function parsePrimary() {
    const token = peek();
    if (token === "(") {
      take();
      const inner = parseOr();
      if (take() !== ")") throw new Error(`Unbalanced parentheses in licence expression: ${expression}`);
      return inner;
    }
    if (!token || token === "AND" || token === "OR" || token === ")") {
      throw new Error(`Unrecognised licence expression syntax: ${expression}`);
    }
    take();
    return { type: "id", id: token };
  }

  function parseAnd() {
    let left = parsePrimary();
    while (peek() === "AND") {
      take();
      left = { type: "AND", left, right: parsePrimary() };
    }
    return left;
  }

  function parseOr() {
    let left = parseAnd();
    while (peek() === "OR") {
      take();
      left = { type: "OR", left, right: parseAnd() };
    }
    return left;
  }

  const ast = parseOr();
  if (index !== tokens.length) throw new Error(`Unrecognised licence expression syntax: ${expression}`);
  return ast;
}

function containsOr(ast) {
  if (ast.type === "OR") return true;
  if (ast.type === "AND") return containsOr(ast.left) || containsOr(ast.right);
  return false;
}

function orBranches(ast) {
  if (ast.type === "OR") return [...orBranches(ast.left), ...orBranches(ast.right)];
  return [ast];
}

function identifiersIn(ast) {
  if (ast.type === "id") return [ast.id];
  return [...identifiersIn(ast.left), ...identifiersIn(ast.right)];
}

function identifiersEqual(left, right) {
  return left.toUpperCase() === right.toUpperCase();
}

function isGplFamily(id) {
  return /^(?:AGPL|GPL)(?:-|$)/i.test(id) && !/^LGPL(?:-|$)/i.test(id);
}

function isAgpl(id) {
  return /^AGPL(?:-|$)/i.test(id);
}

function isRestrictive(id) {
  return /^(?:LGPL|MPL|EPL|CDDL|SSPL|BUSL|COMMONS-CLAUSE|CC-BY-NC)/i.test(id);
}

function classifyIdentifier(id, identity, records) {
  if (ALLOWLIST.has(id.toUpperCase())) return { status: "allowlisted", identifier: id };
  if (isGplFamily(id)) {
    const family = isAgpl(id) ? "AGPL" : "GPL";
    throw new Error(`Production dependency ${identity} is available only under ${family} terms (${id}).`);
  }
  const approval = records.approvals?.[identity];
  const approvedId = approval?.identifier || approval?.approved;
  if (approval && identifiersEqual(approvedId || "", id)) {
    return { status: "approved", identifier: id };
  }
  throw new Error(
    `Production dependency ${identity} declares ${id}, which is outside the permissive allowlist and has no recorded approval.`,
  );
}

function evaluateConjunction(ast, identity, records) {
  if (ast.type === "id") return classifyIdentifier(ast.id, identity, records);
  if (ast.type === "OR") {
    throw new Error(`Production dependency ${identity} declares a nested dual-licence expression that cannot be resolved.`);
  }
  evaluateConjunction(ast.left, identity, records);
  evaluateConjunction(ast.right, identity, records);
  return { status: "allowlisted", identifier: identifiersIn(ast).join(" AND ") };
}

function findElectedBranch(ast, elected) {
  return orBranches(ast).find((branch) => branch.type === "id" && identifiersEqual(branch.id, elected)) || null;
}

function resolvePackageLicence({ identity, declared, records }) {
  const election = records.elections?.[identity];
  if (election && election.declared !== declared) {
    throw new Error(
      `Recorded licence election for ${identity} declared ${election.declared}, but the installed package declares ${declared}.`,
    );
  }

  const ast = parseExpression(declared);
  if (containsOr(ast)) {
    if (!election || !election.elected) {
      throw new Error(
        `Production dependency ${identity} declares ${declared} but has no recorded compatible licence election.`,
      );
    }
    const branch = findElectedBranch(ast, election.elected);
    if (!branch) {
      throw new Error(
        `Recorded licence election ${election.elected} for ${identity} is not an option in ${declared}.`,
      );
    }
    const resolved = evaluateConjunction(branch, identity, records);
    return {
      identity,
      declared,
      status: "elected",
      elected: election.elected,
      identifier: resolved.identifier,
    };
  }

  if (election) {
    throw new Error(`Recorded licence election for ${identity} is not required for declared ${declared}.`);
  }

  const resolved = evaluateConjunction(ast, identity, records);
  return {
    identity,
    declared,
    status: resolved.status,
    identifier: resolved.identifier,
  };
}

function formatElectionNote(identity, declared, elected) {
  const { name, version } = splitIdentity(identity);
  const shown = displayName(name);
  const unelected = identifiersIn(parseExpression(declared)).filter((id) => !identifiersEqual(id, elected));
  const gplOption = unelected.find((id) => isGplFamily(id));
  const lines = [
    "EXHIBIT BUILDER LICENCE ELECTION",
    "-".repeat(78),
    `Exhibit Builder uses ${shown} ${version} under the ${elected} licence option.`,
  ];
  if (gplOption) {
    lines.push(
      `The ${gplOption} text reproduced below forms part of ${shown}'s upstream`,
      `dual-licence notice; Exhibit Builder does not elect the GPL option for its`,
      `use or distribution of ${shown}.`,
    );
  } else {
    lines.push(
      `The remaining dual-licence text reproduced below forms part of ${shown}'s upstream`,
      `notice; Exhibit Builder does not elect a non-${elected} option for its use or`,
      `distribution of ${shown}.`,
    );
  }
  return lines;
}

module.exports = {
  ALLOWLIST,
  formatElectionNote,
  loadLicenceRecords,
  parseExpression,
  resolvePackageLicence,
};