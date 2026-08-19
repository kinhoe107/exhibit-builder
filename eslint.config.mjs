import tseslint from "typescript-eslint";

export default tseslint.config(
  {
    ignores: ["dist/**", "release/**", "node_modules/**", "build/**", "output/**", "Old Builds/**"],
  },
  {
    files: ["app/**/*.{ts,tsx}", "src/**/*.{ts,tsx}", "vite.config.ts"],
    extends: [...tseslint.configs.recommendedTypeChecked],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "module",
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      "@typescript-eslint/await-thenable": "error",
      "@typescript-eslint/consistent-type-imports": ["warn", { prefer: "type-imports", fixStyle: "inline-type-imports" }],
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "warn",
      "@typescript-eslint/no-non-null-assertion": "warn",
      "@typescript-eslint/no-unsafe-argument": "warn",
      "@typescript-eslint/no-unsafe-assignment": "warn",
      "@typescript-eslint/no-unsafe-call": "warn",
      "@typescript-eslint/no-unsafe-member-access": "warn",
      "@typescript-eslint/no-unsafe-return": "warn",
      "@typescript-eslint/no-unnecessary-type-assertion": "warn",
      "@typescript-eslint/no-unused-expressions": "warn",
      "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
      "@typescript-eslint/require-await": "warn",
      "@typescript-eslint/switch-exhaustiveness-check": "error",
      "prefer-const": "warn",
      "complexity": ["warn", 12],
      "max-depth": ["warn", 3]
    },
  },
  {
    files: [
      "app/lib/bundle-arrangement.ts",
      "app/lib/build-plan.ts",
      "app/lib/build-readiness.ts",
      "app/lib/index-layout.ts",
      "app/lib/template-matter-review.ts",
      "app/lib/template-persistence.ts"
    ],
    rules: {
      "@typescript-eslint/no-explicit-any": "error",
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-unsafe-argument": "error",
      "@typescript-eslint/no-unsafe-assignment": "error",
      "@typescript-eslint/no-unsafe-call": "error",
      "@typescript-eslint/no-unsafe-member-access": "error",
      "@typescript-eslint/no-unsafe-return": "error"
    }
  }
);
