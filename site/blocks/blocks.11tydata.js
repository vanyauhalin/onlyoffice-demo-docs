module.exports = {
  eleventyExcludeFromCollections: true,
  permalink(data) {
    if (data.page.fileSlug !== "main") {
      return false
    }
    return `/${data.page.fileSlug}.${data.page.outputFileExtension}`
  }
}
