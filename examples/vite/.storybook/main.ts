import type { StorybookConfig } from "@storybook/web-components-vite";

const config: StorybookConfig = {
  stories: [
    "../src/demo/**/*.mdx",
    "../src/demo/**/*.stories.ts",
  ],
  addons: ["@storybook/addon-docs"],
  framework: {
    name: "@storybook/web-components-vite",
    options: {},
  },
  async viteFinal(baseConfig) {
    const configWithBase = { ...baseConfig };
    const storybookBasePath = process.env.STORYBOOK_BASE_PATH;

    if (storybookBasePath) {
      configWithBase.base = storybookBasePath.endsWith("/")
        ? storybookBasePath
        : `${storybookBasePath}/`;
    }

    return configWithBase;
  },
};

export default config;
