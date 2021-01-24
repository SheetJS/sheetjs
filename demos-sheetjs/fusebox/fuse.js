const { FuseBox } = require("fuse-box");
const common_opts = {
  homeDir: ".",
  output: "$name.js"
};

const browser_opts = {
  target: "browser",
  natives: {
    Buffer: false,
    stream: false,
    process: false
  },
  ...common_opts
};

const node_opts = {
  target: "node",
  ...common_opts
}

const fuse1 = FuseBox.init(browser_opts);
fuse1.bundle("client").instructions(">sheetjs.ts"); fuse1.run();

const fuse2 = FuseBox.init(node_opts);
fuse2.bundle("server").instructions(">sheetjs.ts"); fuse2.run();
