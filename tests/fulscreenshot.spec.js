const { test } = require("@playwright/test");
const path = require("path");
const fs = require("fs");
const sharp = require("sharp");

const CONFIG = {
  url: "https://luvcarwashnew.sharepoint.com/sites/luvdowntimetracking/SitePages/OfflineAssetWorkOrders.aspx",
  outputDir: "screenshots",
  tempDir: "screenshots/temp",
  timeout: 300000,
  viewportWidth: 1920,
  viewportHeight: 1080,
};

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// Find the element that actually scrolls the page content
async function findScrollContainer(page) {
  return await page.evaluate(() => {
    const doc = document.scrollingElement || document.documentElement;
    if (doc.scrollHeight > doc.clientHeight + 50) {
      return { type: "window" };
    }

    const all = Array.from(document.querySelectorAll("body *"));
    let best = null;
    let bestScrollable = 0;

    for (const el of all) {
      const style = getComputedStyle(el);
      const overflowY = style.overflowY;
      const scrollable = el.scrollHeight - el.clientHeight;

      if (
        (overflowY === "auto" || overflowY === "scroll") &&
        scrollable > 100 &&
        el.clientHeight > 200
      ) {
        if (scrollable > bestScrollable) {
          bestScrollable = scrollable;
          best = el;
        }
      }
    }

    if (best) {
      best.setAttribute("data-pw-main-scroll", "true");
      return {
        type: "element",
        scrollHeight: best.scrollHeight,
        clientHeight: best.clientHeight,
      };
    }

    return { type: "none" };
  });
}

async function getScrollState(page, containerType) {
  return await page.evaluate((type) => {
    if (type === "window") {
      return {
        scrollTop: window.scrollY,
        maxScroll:
          (document.scrollingElement || document.documentElement).scrollHeight -
          window.innerHeight,
      };
    } else {
      const el = document.querySelector('[data-pw-main-scroll="true"]');
      if (!el) return { scrollTop: 0, maxScroll: 0 };
      return {
        scrollTop: el.scrollTop,
        maxScroll: el.scrollHeight - el.clientHeight,
      };
    }
  }, containerType);
}

async function scrollBy(page, containerType, amount) {
  await page.evaluate(
    ({ type, amount }) => {
      if (type === "window") {
        window.scrollBy(0, amount);
      } else {
        const el = document.querySelector('[data-pw-main-scroll="true"]');
        if (el) el.scrollTop += amount;
      }
    },
    { type: containerType, amount },
  );
}

// How many pixels from the top AND left of the viewport are NOT part of
// the scrollable content area (e.g. a page header above it, and/or a
// static sidebar to its left). This region is identical in every
// screenshot, so for screenshots after the first we only keep the pixels
// belonging to the actual scroll container.
async function getContainerOffsets(page, containerType) {
  if (containerType !== "element") return { top: 0, left: 0 };
  return await page.evaluate(() => {
    const el = document.querySelector('[data-pw-main-scroll="true"]');
    if (!el) return { top: 0, left: 0 };
    const rect = el.getBoundingClientRect();
    return { top: Math.round(rect.top), left: Math.round(rect.left) };
  });
}

test("SharePoint Full Screenshot (stitched)", async ({ browser }) => {
  test.setTimeout(CONFIG.timeout);

  ensureDir(CONFIG.outputDir);
  ensureDir(CONFIG.tempDir);

  const context = await browser.newContext({
    storageState: "tests/authState.json",
    viewport: {
      width: CONFIG.viewportWidth,
      height: CONFIG.viewportHeight,
    },
  });

  const page = await context.newPage();

  try {
    console.log("Opening page...");

    await page.goto(CONFIG.url, {
      waitUntil: "domcontentloaded",
      timeout: 60000,
    });

    await page.waitForTimeout(8000);

    try {
      await page.waitForSelector(
        '[data-automationid="ListCell"], .ms-List-cell, [role="grid"], [data-automation-key], table',
        { timeout: 30000 },
      );
      console.log("Content detected.");
    } catch (e) {
      console.log("Content selector not matched, continuing anyway...");
    }

    await page.waitForTimeout(3000);

    const containerInfo = await findScrollContainer(page);
    console.log("Scroll container detected:", containerInfo);

    const containerType =
      containerInfo.type === "none" ? "window" : containerInfo.type;

    if (containerType === "window") {
      await page.evaluate(() => window.scrollTo(0, 0));
    } else {
      await page.evaluate(() => {
        const el = document.querySelector('[data-pw-main-scroll="true"]');
        if (el) el.scrollTop = 0;
      });
    }
    await page.waitForTimeout(1500);

    // Measure the static (non-scrolling) region at the top and left once.
    const { top: containerTopOffset, left: containerLeftOffset } =
      await getContainerOffsets(page, containerType);
    // The actual pixel height of scrollable content visible per screenshot.
    const contentAreaHeight =
      containerType === "element"
        ? containerInfo.clientHeight
        : CONFIG.viewportHeight;
    console.log("Static top offset detected:", containerTopOffset);
    console.log(
      "Static left offset detected (sidebar width):",
      containerLeftOffset,
    );
    console.log("Scrollable content area height:", contentAreaHeight);

    // CRITICAL: never scroll by more than the visible content height, or a
    // strip of content between screenshots is skipped entirely (never
    // captured by any screenshot) — this is what caused missing rows.
    // Reserve a buffer so consecutive screenshots always share real overlap,
    // which the pixel-matching step needs to find the correct seam.
    const overlapBuffer = 60;
    const scrollStep = Math.max(200, contentAreaHeight - overlapBuffer);
    console.log("Scroll step per iteration:", scrollStep);

    const screenshotPaths = [];
    // cropTop[i] = how many pixels to crop off the TOP of screenshot i
    // before stitching (removes the repeated sticky header + overlap with
    // the previous screenshot).
    const cropTop = [];
    let step = 0;
    let stableCount = 0;
    const maxSteps = 100;

    let prevScrollTop = 0;

    while (step < maxSteps) {
      const shotPath = path.join(
        CONFIG.tempDir,
        `step-${String(step).padStart(3, "0")}.png`,
      );
      await page.screenshot({ path: shotPath });

      const stateBefore = await getScrollState(page, containerType);
      console.log(
        `Step ${step}: scrollTop=${stateBefore.scrollTop}, maxScroll=${stateBefore.maxScroll}`,
      );

      if (step === 0) {
        cropTop.push(0); // first screenshot: keep everything
      } else {
        // Actual pixels scrolled since previous screenshot (not assumed!)
        const actualDelta = stateBefore.scrollTop - prevScrollTop;
        // Content already shown in the previous shot, measured against the
        // REAL scrollable content area height (not the full viewport).
        const overlapWithPrevious = Math.max(
          0,
          contentAreaHeight - actualDelta,
        );
        // Crop out the static top region (repeated header/breadcrumb) AND
        // the overlapping content beneath it.
        cropTop.push(containerTopOffset + overlapWithPrevious);
      }
      prevScrollTop = stateBefore.scrollTop;
      screenshotPaths.push(shotPath);

      if (stateBefore.scrollTop >= stateBefore.maxScroll - 5) {
        console.log("Reached bottom, stopping.");
        break;
      }

      await scrollBy(page, containerType, scrollStep);
      await page.mouse.wheel(0, scrollStep);
      await page.waitForTimeout(2800);

      const stateAfter = await getScrollState(page, containerType);

      if (stateAfter.scrollTop === stateBefore.scrollTop) {
        stableCount++;
        if (stableCount >= 2) {
          console.log("Scroll position not changing, stopping.");
          break;
        }
      } else {
        stableCount = 0;
      }

      step++;
    }

    console.log(
      `Captured ${screenshotPaths.length} viewport screenshots. Stitching...`,
    );

    // NOTE: pixel-based overlap refinement was removed. It compared rendered
    // pixels between screenshots to fine-tune the crop, but browsers can
    // render the SAME static content with tiny anti-aliasing differences
    // between two separate screenshots, which was introducing noise and
    // causing wrong matches (duplicated or skipped rows). The scrollTop
    // values from the browser are exact (including sub-pixel precision), so
    // the math-based crop below is fully deterministic and doesn't need
    // pixel verification.
    for (let i = 1; i < screenshotPaths.length; i++) {
      console.log(`Step ${i}: crop (math) = ${Math.round(cropTop[i])}px`);
    }

    const images = await Promise.all(
      screenshotPaths.map(async (p, i) => {
        const meta = await sharp(p).metadata();
        const cropT = Math.round(cropTop[i]);
        // Only the first screenshot keeps the sidebar; later ones crop it
        // out on the left since it's static and already captured once.
        const cropL = i === 0 ? 0 : containerLeftOffset;
        const needsCrop = cropT > 0 || cropL > 0;

        if (needsCrop && cropT < meta.height && cropL < meta.width) {
          const croppedPath = p.replace(".png", "-cropped.png");
          await sharp(p)
            .extract({
              left: cropL,
              top: cropT,
              width: meta.width - cropL,
              height: meta.height - cropT,
            })
            .toFile(croppedPath);
          const croppedMeta = await sharp(croppedPath).metadata();
          return {
            path: croppedPath,
            width: croppedMeta.width,
            height: croppedMeta.height,
            left: cropL, // where this strip belongs horizontally in the final image
          };
        }
        return { path: p, width: meta.width, height: meta.height, left: 0 };
      }),
    );

    const totalWidth = images[0].width;
    const totalHeight = images.reduce((sum, img) => sum + img.height, 0);

    let yOffset = 0;
    const compositeInputs = images.map((img) => {
      const input = { input: img.path, top: yOffset, left: img.left };
      yOffset += img.height;
      return input;
    });

    const finalPath = path.join(
      CONFIG.outputDir,
      "SharePoint-FullPage-Stitched.png",
    );

    await sharp({
      create: {
        width: totalWidth,
        height: totalHeight,
        channels: 3,
        background: { r: 255, g: 255, b: 255 },
      },
    })
      .composite(compositeInputs)
      .png()
      .toFile(finalPath);

    console.log("Stitched screenshot saved:", path.resolve(finalPath));

    fs.rmSync(CONFIG.tempDir, { recursive: true, force: true });

    console.log("Finished successfully.");
    console.log("Saved to:", path.resolve(CONFIG.outputDir));
  } catch (err) {
    console.error(err);
    await page.screenshot({
      path: path.join(CONFIG.outputDir, "Error.png"),
      fullPage: true,
    });
  } finally {
    await context.close();
  }
});
