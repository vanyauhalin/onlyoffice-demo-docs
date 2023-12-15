const { basename } = require("node:path")
const definitions = require("@onlyoffice-demo-docs/document-builder/data/cell.json")

function setup() {
  const data = {
    layout: "layouts/member/member.njk",
    pagination: {
      data: "items",
      size: 1,
      addAllPagesToCollections: true
    },
    items: [],
    permalink(data) {
      const { memberof, name } = data.pagination.items[0]
      return `/document-builder/javascript-sdk/spreadsheet/${memberof}/${name}/index.html`
    },
    eleventyComputed: {
      title(data) {
        return basename(data.page.url)
      },
      currentName(data) {
        return basename(data.page.url)
      }
    }
  }

  definitions.forEach((d) => {
    switch (d.kind) {
      case "member":
        if (d.name === undefined || d.memberof === undefined) {
          break
        }

        // todo: SetMinorVerticalGridlines includes twice (bug)
        if (d.name === "SetMinorVerticalGridlines") {
          break
        }

        const t = {}
        t.name = d.name
        t.memberof = d.memberof
        t.description = d.description || ""
        t.signatures = ["console.log(\"I'll do it tomorrow\")"]
        t.parameters = (d.params || []).map((p) => {
          return {
            name: p.name,
            type: normalizeType(p.type.names),
            required: !!p.optional,
            description: p.description,
            default: p.defaultvalue
            // example:
          }
        })
        t.returns = (d.returns || []).map((p) => {
          return {
            type: normalizeType(p.type.names),
          }
        })
        t.examples = d.examples

        insert(data.items, t, "name")
    }
  })

  return data
}

/**
 * @param {string[]} a
 * @returns {string}
 */
function normalizeType(a) {
  const t = a.map((n) => n.replace(".", "")).join(" | ")
  switch (t) {
    case "Array":
      return "[]"
    case "Boolean":
      return "boolean"
    case "Number":
      return "number"
    case "String":
      return "string"
    default:
      return t
  }
}

/**
 * @typedef {Object} Definition
 * @property {string} name
 */

/**
 * @param {Definition[]} a
 * @param {Definition} d
 * @param {keyof Definition} k
 * @returns {Definition[]}
 */
function insert(a, d, k) {
  let s = 0
  let e = a.length - 1
  while (s <= e) {
    let m = Math.floor((s + e) / 2)
    if (a[m][k] === d[k]) {
      a.splice(m, 0, d)
      return a
    }
    if (a[m][k] < d[k]) {
      s = m + 1
    } else {
      e = m - 1
    }
  }
  a.splice(s, 0, d)
  return a
}

module.exports = setup()
