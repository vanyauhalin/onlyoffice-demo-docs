declare global {
  interface Window {
    BTSnippet: typeof BTSnippet
  }

  interface HTMLElementTagNameMap {
    "bt-snippet": BTSnippet
  }
}
