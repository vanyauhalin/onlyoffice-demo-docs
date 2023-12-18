const { basename } = require("node:path")
const definitions = require("@onlyoffice-demo-docs/document-builder/data/cell.json")

function setup() {
  const data = {
    layout: "class/class.webc",
    pagination: {
      data: "items",
      size: 1,
      addAllPagesToCollections: true
    },
    items: [],
    permalink(data) {
      const { name } = data.pagination.items[0]
      return `/document-builder/javascript-sdk/spreadsheet/${name}/index.html`
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
      case "class":
        if (d.name === undefined) {
          break
        }

        const t = {}
        t.name = d.name
        t.description = d.description
        t.signatures = ["console.log(\"I'll do it tomorrow\")"]
        t.properties = (d.properties || []).map((p) => {
          const o = {
            name: p.name
          }
          if (p.description) {
            o.description = p.description
          }
          if (p.type) {
            o.type = normalizeType(p.type.names)
          }
          return o
        })

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
  const t = a.map((n) => n.replaceAll(".", "")).join(" | ")
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
