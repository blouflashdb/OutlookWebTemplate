import antfu from '@antfu/eslint-config'

export default antfu({
  vue: true,
  typescript: {
    // Enables type-aware rules. The solution tsconfig resolves each file to its project via references.
    tsconfigPath: 'tsconfig.json',
    // Matches `erasableSyntaxOnly` from @tsconfig/node-ts.
    erasableOnly: true,
  },
  pnpm: true,
})
