// Lets plain TypeScript tooling (type-aware ESLint) import .vue files.
// vue-tsc resolves the actual components and ignores this fallback.
declare module '*.vue' {
  import type { DefineComponent } from 'vue'

  const component: DefineComponent
  export default component
}
