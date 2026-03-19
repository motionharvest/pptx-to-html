import type { Preview } from "@storybook/web-components-vite";

const preview: Preview = {
  parameters: {
    layout: "padded",
    controls: {
      expanded: true,
    },
    options: {
      storySort: {
        order: ["Library"],
      },
    },
  },
};

export default preview;
