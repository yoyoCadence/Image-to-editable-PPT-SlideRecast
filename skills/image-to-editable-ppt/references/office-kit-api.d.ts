//#region src/api/units.d.ts
declare const emuBrand: unique symbol;
type Emu = number & {
  readonly [emuBrand]: 'Emu';
};
declare const inches: (n: number) => Emu;
declare const cm: (n: number) => Emu;
declare const mm: (n: number) => Emu;
declare const pt: (n: number) => Emu;
declare const emu: (n: number) => Emu;
//#endregion
//#region src/internal/opc/part-name.d.ts
type PartName = string & {
  readonly __brand: 'PartName';
};
//#endregion
//#region src/internal/opc/content-types.d.ts
interface ContentTypeDefault {
  /** Lowercase, no leading dot, e.g. `xml`. */
  extension: string;
  contentType: string;
}
interface ContentTypeOverride {
  partName: PartName;
  contentType: string;
}
interface ContentTypes {
  defaults: ContentTypeDefault[];
  overrides: ContentTypeOverride[];
}
//#endregion
//#region src/internal/opc/rels.d.ts
type TargetMode = 'Internal' | 'External';
interface Relationship {
  id: string;
  type: string;
  target: string;
  /** Defaults to `'Internal'` when omitted from the XML. */
  targetMode: TargetMode;
}
interface Relationships {
  items: Relationship[];
}
//#endregion
//#region src/internal/opc/image-format.d.ts
type ImageFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'tiff' | 'webp' | 'svg';
//#endregion
//#region src/internal/opc/media-format.d.ts
type VideoFormat = 'mp4' | 'm4v' | 'mov' | 'webm' | 'avi' | 'wmv';
type AudioFormat = 'mp3' | 'wav' | 'm4a' | 'ogg' | 'wma';
//#endregion
//#region src/internal/parts/package.d.ts
/** A single OPC part: name, content type, and bytes. */
interface Part {
  readonly name: PartName;
  contentType: string;
  data: Uint8Array;
}
/**
 * In-memory model of an OPC package. Construct via `OpcPackage.load(bytes)`
 * to read an existing `.pptx`, or via `OpcPackage.empty()` to start from
 * scratch.
 */
declare class OpcPackage {
  /**
   * Ordered list of parts. The order is the order parts will be written to
   * the ZIP central directory on `save()`. Direct mutation is allowed for
   * callers that need to reorder, but `addPart` / `removePart` are the
   * supported entry points.
   */
  readonly parts: Part[];
  /** Mutable Content_Types map. Direct mutation is allowed. */
  contentTypes: ContentTypes;
  private constructor();
  /**
   * Builds an empty package containing only the two default content-type
   * entries (xml, rels). No `.rels` or content parts yet — that's the caller's
   * job. PresentationML scaffolding will populate the typical PPTX skeleton.
   */
  static empty(): OpcPackage;
  /**
   * Loads a `.pptx` file from raw bytes. Parses `[Content_Types].xml`
   * eagerly; leaves part bodies as raw `Uint8Array`s for higher layers to
   * interpret.
   *
   * Throws if the archive is not a valid ZIP, or if `[Content_Types].xml` is
   * missing or malformed.
   */
  static load(bytes: Uint8Array): OpcPackage;
  /** Serialize the package to a `.pptx` byte stream. */
  save(): Uint8Array;
  /**
   * Looks up a part by name (case-insensitive per OPC §9.1.1.7).
   */
  getPart(name: PartName): Part | null;
  /**
   * Adds a new part. Throws if a part with the same name (case-insensitive)
   * already exists. The caller is responsible for ensuring the content type
   * is resolvable — either by extension default or by a registered override.
   */
  addPart(name: PartName, contentType: string, data: Uint8Array): Part;
  /**
   * Removes a part by name. Returns true if a part was removed; false if no
   * part with that name was present. Also drops any matching `Override`
   * entry from Content_Types; defaults are left untouched.
   */
  removePart(name: PartName): boolean;
  /**
   * Reads (or parses, lazily) the relationships for the given part. Returns
   * `null` if the part has no sibling `.rels` part — this is normal for
   * content parts with no outgoing relationships.
   *
   * The returned `Relationships` is a fresh snapshot. Mutate it as needed,
   * then call `setRels(forPart, rels)` to persist the change back into the
   * package.
   */
  getRels(forPart: PartName): Relationships | null;
  /**
   * Writes relationships for the given part. Creates the sibling `.rels`
   * part if it didn't exist; replaces its bytes if it did. The relationships
   * part is given the standard relationships content type via the `rels`
   * extension default.
   */
  setRels(forPart: PartName, rels: Relationships): void;
  /**
   * Convenience: read the package-root relationships (`/_rels/.rels`).
   * These point at the top-level parts — typically the presentation.xml,
   * core properties, and extended properties.
   */
  rootRels(): Relationships | null;
  /** Convenience: write package-root relationships. */
  setRootRels(rels: Relationships): void;
}
//#endregion
//#region src/internal/xml/ast.d.ts
interface QName {
  /** Empty string when no prefix is used (default namespace or unqualified). */
  readonly prefix: string;
  readonly localName: string;
  /** Empty string when the name has no namespace binding. */
  readonly namespaceURI: string;
}
interface XmlAttr {
  readonly name: QName;
  readonly value: string;
}
interface XmlElement {
  readonly kind: 'element';
  name: QName;
  attrs: XmlAttr[];
  /**
   * Prefix → URI declarations that appeared as `xmlns` / `xmlns:foo`
   * attributes on this element. The serializer re-emits them as attributes
   * (NOT inside `attrs`, since that array is for non-namespace attributes
   * only). An empty-string prefix entry is a default-namespace declaration.
   */
  prefixDecls: Map<string, string>;
  children: XmlNode[];
}
interface XmlText {
  readonly kind: 'text';
  data: string;
}
interface XmlCData {
  readonly kind: 'cdata';
  data: string;
}
interface XmlComment {
  readonly kind: 'comment';
  data: string;
}
interface XmlPI {
  readonly kind: 'pi';
  target: string;
  data: string;
}
type XmlNode = XmlElement | XmlText | XmlCData | XmlComment | XmlPI;
interface XmlDeclaration {
  readonly version: string;
  readonly encoding?: string;
  readonly standalone?: 'yes' | 'no';
}
interface XmlDocument {
  readonly kind: 'document';
  decl: XmlDeclaration | null;
  /** Comments / processing instructions appearing before the root element. */
  prolog: ReadonlyArray<XmlComment | XmlPI>;
  root: XmlElement;
  /** Comments / processing instructions appearing after the root element. */
  epilog: ReadonlyArray<XmlComment | XmlPI>;
}
//#endregion
//#region src/internal/enum-values.d.ts
declare const PLACEHOLDER_TYPES: readonly ["title", "body", "ctrTitle", "subTitle", "dt", "sldNum", "ftr", "hdr", "obj", "chart", "tbl", "clipArt", "dgm", "media", "sldImg", "pic"];
declare const CAMERA_PRESETS: readonly ["legacyObliqueTopLeft", "legacyObliqueTop", "legacyObliqueTopRight", "legacyObliqueLeft", "legacyObliqueFront", "legacyObliqueRight", "legacyObliqueBottomLeft", "legacyObliqueBottom", "legacyObliqueBottomRight", "legacyPerspectiveTopLeft", "legacyPerspectiveTop", "legacyPerspectiveTopRight", "legacyPerspectiveLeft", "legacyPerspectiveFront", "legacyPerspectiveRight", "legacyPerspectiveBottomLeft", "legacyPerspectiveBottom", "legacyPerspectiveBottomRight", "orthographicFront", "isometricTopUp", "isometricTopDown", "isometricBottomUp", "isometricBottomDown", "isometricLeftUp", "isometricLeftDown", "isometricRightUp", "isometricRightDown", "isometricOffAxis1Left", "isometricOffAxis1Right", "isometricOffAxis1Top", "isometricOffAxis2Left", "isometricOffAxis2Right", "isometricOffAxis2Top", "isometricOffAxis3Left", "isometricOffAxis3Right", "isometricOffAxis3Bottom", "isometricOffAxis4Left", "isometricOffAxis4Right", "isometricOffAxis4Bottom", "obliqueTopLeft", "obliqueTop", "obliqueTopRight", "obliqueLeft", "obliqueRight", "obliqueBottomLeft", "obliqueBottom", "obliqueBottomRight", "perspectiveFront", "perspectiveLeft", "perspectiveRight", "perspectiveAbove", "perspectiveBelow", "perspectiveAboveLeftFacing", "perspectiveAboveRightFacing", "perspectiveContrastingLeftFacing", "perspectiveContrastingRightFacing", "perspectiveHeroicLeftFacing", "perspectiveHeroicRightFacing", "perspectiveHeroicExtremeLeftFacing", "perspectiveHeroicExtremeRightFacing", "perspectiveRelaxed", "perspectiveRelaxedModerately"];
declare const LIGHT_RIG_TYPES: readonly ["legacyFlat1", "legacyFlat2", "legacyFlat3", "legacyFlat4", "legacyNormal1", "legacyNormal2", "legacyNormal3", "legacyNormal4", "legacyHarsh1", "legacyHarsh2", "legacyHarsh3", "legacyHarsh4", "threePt", "balanced", "soft", "harsh", "flood", "contrasting", "morning", "sunrise", "sunset", "chilly", "freezing", "flat", "twoPt", "glow", "brightRoom"];
declare const LIGHT_RIG_DIRECTIONS: readonly ["tl", "t", "tr", "l", "r", "bl", "b", "br"];
declare const BEVEL_PRESETS: readonly ["relaxedInset", "circle", "slope", "cross", "angle", "softRound", "convex", "coolSlant", "divot", "riblet", "hardEdge", "artDeco"];
declare const PRESET_MATERIALS: readonly ["legacyMatte", "legacyPlastic", "legacyMetal", "legacyWireframe", "matte", "plastic", "metal", "warmMatte", "translucentPowder", "powder", "dkEdge", "softEdge", "clear", "flat", "softmetal"];
//#endregion
//#region src/internal/presentationml/slide-part.d.ts
type ShapeKind = 'shape' | 'picture' | 'group' | 'graphicFrame' | 'connector';
/**
 * Placeholder token from `<p:ph type="...">` (ECMA-376 `ST_PlaceholderType`).
 * Placeholder lookups take this closed set, while `SlideShape.placeholderType`
 * stays `string | null`: a loaded deck can carry a value outside it.
 */
type PlaceholderType = (typeof PLACEHOLDER_TYPES)[number];
interface SlideShape {
  readonly kind: ShapeKind;
  /** OOXML internal numeric id, unique within the slide's shape tree. */
  readonly id: number;
  /** Human-readable name (`Title 1`, `Content Placeholder 2`, ...). */
  readonly name: string;
  /**
   * Placeholder type (`title`, `body`, `ctrTitle`, ...) when the shape is a
   * placeholder; `null` for regular shapes and pictures. From `p:ph/@type`.
   */
  readonly placeholderType: string | null;
  /**
   * Placeholder index, used to bind the slide-level shape to its same-idx
   * counterpart on the slide layout / master.
   */
  readonly placeholderIdx: number | null;
  /** All visible text concatenated. `''` if the shape has no text body. */
  readonly text: string;
  /**
   * The underlying XML element for this shape. Higher layers walk this when
   * they need formatting / geometry information not yet promoted to the
   * typed view.
   */
  readonly element: XmlElement;
}
interface SlidePart {
  readonly shapes: ReadonlyArray<SlideShape>;
  readonly root: XmlElement;
}
//#endregion
//#region src/internal/presentationml/slide-layout-part.d.ts
/**
 * Layout type token per ECMA-376 §19.7.15 `ST_SlideLayoutType`. Listed for
 * autocomplete; layouts in real templates sometimes use vendor-specific
 * tokens not in this enum, so the field stays a plain string at the API.
 */
type SlideLayoutType = 'title' | 'tx' | 'twoColTx' | 'tbl' | 'txAndChart' | 'chartAndTx' | 'dgm' | 'chart' | 'txAndClipArt' | 'clipArtAndTx' | 'titleOnly' | 'blank' | 'txAndObj' | 'objAndTx' | 'objOnly' | 'obj' | 'txAndMedia' | 'mediaAndTx' | 'objOverTx' | 'txOverObj' | 'txAndTwoObj' | 'twoObjAndTx' | 'twoObjOverTx' | 'fourObj' | 'vertTx' | 'clipArtAndVertTx' | 'vertTitleAndTx' | 'vertTitleAndTxOverChart' | 'twoObj' | 'objAndTwoObj' | 'twoObjAndObj' | 'cust' | 'secHead' | 'twoTxTwoObj' | 'objTx' | 'picTx';
interface SlideLayoutPart {
  /** Human-visible layout name from `<p:cSld name="...">`. */
  readonly name: string;
  /**
   * Layout type token. `null` when the attribute is absent — the spec
   * default in that case is `cust` (custom).
   */
  readonly layoutType: string | null;
  /** Shapes on the layout in document order, groups flattened. */
  readonly shapes: ReadonlyArray<SlideShape>;
  readonly root: XmlElement;
}
//#endregion
//#region src/internal/presentationml/shape-builder.d.ts
/**
 * Common ECMA-376 `ST_ShapeType` preset tokens. The schema accepts ~180
 * presets; the union below covers the ones authoring scripts reach for
 * most often. Pass any string to `addShape({ preset })` to use a token
 * not enumerated here.
 */
type PresetShape = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rtTriangle' | 'diamond' | 'parallelogram' | 'trapezoid' | 'pentagon' | 'hexagon' | 'heptagon' | 'octagon' | 'decagon' | 'star4' | 'star5' | 'star6' | 'star7' | 'star8' | 'star10' | 'star12' | 'star16' | 'star24' | 'star32' | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'leftRightArrow' | 'upDownArrow' | 'bentArrow' | 'curvedRightArrow' | 'curvedLeftArrow' | 'curvedUpArrow' | 'curvedDownArrow' | 'cloud' | 'heart' | 'lightningBolt' | 'sun' | 'moon' | 'bracketPair' | 'bracePair' | 'leftBracket' | 'rightBracket' | 'leftBrace' | 'rightBrace' | 'cube' | 'can' | 'donut' | 'noSmoking' | 'plus' | 'mathPlus' | 'mathMinus' | 'mathMultiply' | 'mathDivide' | 'mathEqual' | 'mathNotEqual';
//#endregion
//#region src/internal/drawingml/color-transforms.d.ts
/** Ordered DrawingML color adjustments. Percentages are fractions; hue angles are degrees. */
type ColorTransform = {
  readonly kind: 'comp' | 'inv' | 'gray' | 'gamma' | 'invGamma';
} | {
  readonly kind: 'tint' | 'shade' | 'alpha' | 'alphaOff' | 'alphaMod' | 'hue' | 'hueOff' | 'hueMod' | 'sat' | 'satOff' | 'satMod' | 'lum' | 'lumOff' | 'lumMod' | 'red' | 'redOff' | 'redMod' | 'green' | 'greenOff' | 'greenMod' | 'blue' | 'blueOff' | 'blueMod';
  readonly value: number;
};
//#endregion
//#region src/internal/drawingml/color.d.ts
declare const SCHEME_COLOR_TOKENS: readonly ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink", "phClr", "lt1", "dk1", "lt2", "dk2"];
/** A theme color slot (`<a:schemeClr val="..."/>`). */
type SchemeColorToken = (typeof SCHEME_COLOR_TOKENS)[number];
/**
 * An sRGB color, `#RRGGBB` or the 3-digit `#RGB` shorthand.
 *
 * TypeScript cannot spell "six hex digits", so a malformed body such as
 * `'#zzzzzz'` still type-checks and is rejected at run time. Requiring the
 * `#` is what stops the union collapsing to `string` and taking every other
 * color mistake down with it.
 */
type HexColor = `#${string}`;
/**
 * Any color an authoring call accepts. Contexts that cannot resolve a theme
 * slot — chart series fills, which must carry a concrete sRGB value — take
 * `HexColor` instead.
 */
type Color = HexColor | SchemeColorToken | `scheme:${SchemeColorToken}`;
/**
 * Narrows a color read back from a deck to one that can be written again,
 * or `null` when it is not a color this library can emit.
 *
 * The readers surface `<a:schemeClr val>` verbatim, so a round-trip
 * (`setX(getX(...))`) has to pass through a check somewhere; doing it here
 * keeps that single `Color` assertion backed by the same parser that
 * `buildColorElement` uses.
 */
declare const asColor: (value: string) => Color | null;
//#endregion
//#region src/internal/drawingml/effects.d.ts
interface ShadowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. Defaults to black. */
  readonly color?: Color;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, ...), written
   * as its children — the same field gradient stops take. Requires `color`.
   * An `opacity` that differs from the one these transforms state wins.
   */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Edge blur in EMU. Defaults to 50800 (4pt). */
  readonly blurEmu?: number;
  /** Offset distance in EMU. Defaults to 38100 (3pt). */
  readonly offsetEmu?: number;
  /**
   * Direction in degrees, measured clockwise from the right (3 o'clock).
   * Defaults to 45° (down-right).
   */
  readonly angleDeg?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
  /** Transform anchor; defaults to top left (`tl`) for authored shadows. */
  readonly alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  /** Whether the shadow rotates with its shape; defaults to false. */
  readonly rotateWithShape?: boolean;
  /**
   * Horizontal scale (`sx`); defaults to one. The reference desktop app's Size field sets
   * both scales, and its perspective presets flatten the shadow with `sy`.
   */
  readonly scaleX?: number;
  /** Vertical scale (`sy`); defaults to one. Negative values mirror vertically. */
  readonly scaleY?: number;
  /** Horizontal skew (`kx`) in degrees, strictly between -90 and 90. */
  readonly skewX?: number;
  /** Vertical skew (`ky`) in degrees, strictly between -90 and 90. */
  readonly skewY?: number;
}
/** Character or shape inner-shadow parameters from `<a:innerShdw>`. */
interface InnerShadowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. Defaults to black. */
  readonly color?: Color;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, ...), written
   * as its children — the same field gradient stops take. Requires `color`.
   * An `opacity` that differs from the one these transforms state wins.
   */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Edge blur in EMU. Defaults to 50800 (4pt). */
  readonly blurEmu?: number;
  /** Offset distance in EMU. Defaults to 38100 (3pt). */
  readonly offsetEmu?: number;
  /** Direction in degrees, measured clockwise from the right. Defaults to 45°. */
  readonly angleDeg?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}
interface GlowOptions {
  /** `#RRGGBB`, bare `RRGGBB`, or scheme token. */
  readonly color: Color;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, ...), written
   * as its children — the same field gradient stops take. Requires `color`.
   * An `opacity` that differs from the one these transforms state wins.
   */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Glow radius in EMU. Defaults to 63500 (5pt). */
  readonly radiusEmu?: number;
  /** Opacity (0–1). Defaults to fully opaque. */
  readonly opacity?: number;
}
/** Character reflection parameters from DrawingML's `<a:reflection>`. */
interface ReflectionOptions {
  /** Blur radius in EMU; defaults to zero. */
  readonly blurEmu?: number;
  /** Offset distance in EMU; defaults to zero. */
  readonly offsetEmu?: number;
  /** Offset direction in clockwise degrees from right; defaults to zero. */
  readonly angleDeg?: number;
  /** Ending opacity (0–1); defaults to zero (`endA`). */
  readonly opacity?: number;
  /** Starting opacity (0–1); defaults to one. */
  readonly startOpacity?: number;
  /** Starting fade position (0–1); defaults to zero. */
  readonly startPosition?: number;
  /** Ending fade position (0–1); defaults to one. */
  readonly endPosition?: number;
  /** Fade direction in clockwise degrees from right; defaults to 90. */
  readonly fadeDirection?: number;
  /** Horizontal scale; defaults to one. Negative values mirror horizontally. */
  readonly scaleX?: number;
  /** Vertical scale; defaults to one. Negative values mirror vertically. */
  readonly scaleY?: number;
  /** Horizontal skew in degrees, strictly between -90 and 90. */
  readonly skewX?: number;
  /** Vertical skew in degrees, strictly between -90 and 90. */
  readonly skewY?: number;
  /** Transform anchor; defaults to bottom center (`b`). */
  readonly alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  /** Whether the effect rotates with its shape; defaults to true. */
  readonly rotateWithShape?: boolean;
}
//#endregion
//#region src/internal/drawingml/fill.d.ts
/** One stop in a linear gradient. */
interface GradientStop {
  /** Position on the gradient axis, 0-1. */
  readonly offset: number;
  /** `#RRGGBB`, the `#RGB` shorthand, or a theme color token. */
  readonly color: Color;
  /** Ordered imported color adjustments. Brightness and opacity override their corresponding transforms when changed. */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Opacity from 0 (transparent) to 1 (opaque). */
  readonly opacity?: number;
  /** The reference desktop app's brightness from -1 (black) to 1 (white); 0 leaves the color unchanged. */
  readonly brightness?: number;
}
interface GradientFillOptions {
  /** Two or more color stops along the axis. */
  readonly stops: ReadonlyArray<GradientStop>;
  /**
   * Gradient axis angle, in degrees. ECMA-376 measures clockwise from
   * the right (3 o'clock); `0` is a left-to-right gradient, `90` is
   * top-to-bottom, `180` right-to-left, `270` bottom-to-top. Defaults
   * to `90` (top → bottom). Only meaningful for linear gradients.
   */
  readonly angleDeg?: number;
  /** Rotate the gradient with the shape. Defaults to true. */
  readonly rotateWithShape?: boolean;
  /** Scale a linear gradient with the shape's aspect ratio. Defaults to false. */
  readonly scaled?: boolean;
  /**
   * Non-linear gradient path. `circle` paints concentric circles,
   * `rect` paints nested rectangles, `shape` follows the shape's
   * outline. Absent / `'linear'` → linear gradient using `angleDeg`.
   * Mirrors ECMA-376 §20.1.8.33 `<a:path path="…"/>`.
   */
  readonly path?: 'linear' | 'circle' | 'rect' | 'shape';
  /**
   * Focus rectangle for non-linear gradients, expressed as fractional insets
   * from each corresponding edge (1 = 100%). Four 0.5 insets describe
   * a point at the center; left/top 1 and right/bottom 0 describe the
   * bottom-right corner. Mirrors `<a:fillToRect>`.
   */
  readonly focus?: {
    readonly left: number;
    readonly top: number;
    readonly right: number;
    readonly bottom: number;
  };
  /**
   * Gradient tile bounds as fractional insets from the shape's edges.
   * Negative values extend the tile beyond that edge. The reference desktop app on Mac uses
   * right/bottom -1 for its bottom-right radial direction. Mirrors `<a:tileRect>`.
   */
  readonly tileRect?: GradientFillOptions['focus'];
}
/**
 * A gradient read back from a deck. Stop colors widen to `string`: the
 * readers surface a scheme token verbatim, including one outside the theme,
 * rather than dropping it.
 */
type ReadGradientStop = Omit<GradientStop, 'color'> & {
  readonly color: string; /** Theme- and transform-resolved color, supplied by the effective shape reader. */
  readonly resolvedColor?: string;
};
type ReadGradientFill = Omit<GradientFillOptions, 'stops'> & {
  readonly stops: ReadonlyArray<ReadGradientStop>;
};
/** Every `ST_PresetPatternVal` token (ECMA-376 dml-main.xsd). */
declare const PATTERN_PRESETS: readonly ["pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70", "pct75", "pct80", "pct90", "horz", "vert", "ltHorz", "ltVert", "dkHorz", "dkVert", "narHorz", "narVert", "dashHorz", "dashVert", "cross", "dnDiag", "upDiag", "ltDnDiag", "ltUpDiag", "dkDnDiag", "dkUpDiag", "wdDnDiag", "wdUpDiag", "dashDnDiag", "dashUpDiag", "diagCross", "smCheck", "lgCheck", "smGrid", "lgGrid", "dotGrid", "smConfetti", "lgConfetti", "horzBrick", "diagBrick", "solidDmnd", "openDmnd", "dotDmnd", "plaid", "sphere", "weave", "divot", "shingle", "wave", "trellis", "zigZag"];
/** One of ECMA-376's `ST_PresetPatternVal` tokens (`pct50`, `dkUpDiag`, ...). */
type PatternPreset = (typeof PATTERN_PRESETS)[number];
interface PatternFillOptions {
  /** Preset pattern token, e.g. `'pct50'`, `'dkUpDiag'`, `'wave'`. */
  readonly preset: PatternPreset;
  /** Foreground (pattern stroke) color. `#RRGGBB`, bare `RRGGBB`, or scheme token. */
  readonly foreground: string;
  /** Background (fill behind the pattern) color. */
  readonly background: string;
  readonly foregroundTransforms?: readonly ColorTransform[];
  readonly backgroundTransforms?: readonly ColorTransform[];
}
//#endregion
//#region src/internal/drawingml/stroke.d.ts
/**
 * A non-solid line paint — the fill choice of `CT_LineProperties` other than
 * `<a:solidFill>` (which `color` spells) and `<a:noFill>`. Only gradients are
 * modelled; the reference desktop app's Format pane offers no other kind for a line.
 */
type LineFill = {
  readonly kind: 'gradient';
} & GradientFillOptions;
/** A line paint read back from a deck; stop colors widen to `string`. */
type ReadLineFill = {
  readonly kind: 'gradient';
} & ReadGradientFill;
/**
 * ECMA-376 §20.1.10.49 `ST_PresetLineDashVal` tokens. The reference desktop app's
 * "Dash type" dropdown maps to these.
 */
type LineDash = 'solid' | 'dot' | 'dash' | 'lgDash' | 'dashDot' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot';
/** ECMA-376 §20.1.10.39 `ST_LineEndType`. */
type LineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
/** ECMA-376 `ST_LineEndWidth` / `ST_LineEndLength`. */
type LineEndSize = 'sm' | 'med' | 'lg';
interface ArrowOptions {
  readonly type: LineEndType;
  readonly width?: LineEndSize;
  readonly length?: LineEndSize;
}
/**
 * The reference desktop app's Sketched style presets ([MS-ODRAWXML] §2.38,
 * `EG_LineSketchType`), named after their `ask:lineSketch*` elements.
 */
type LineSketch = 'curved' | 'freehand' | 'scribble';
//#endregion
//#region src/internal/drawingml/text-format.d.ts
interface TextFormat {
  /** Latin font family (`Calibri`, `Arial`, ...). Sets `<a:latin>`. */
  font?: string;
  /**
   * East Asian font family (`游明朝`, `メイリオ`, ...). Sets `<a:ea>`.
   * Renderers pick this typeface for CJK glyphs independently of `font`
   * (which only governs Latin glyphs) — set both when a run mixes Latin
   * and CJK text and needs a consistent look across the whole run.
   */
  fontEastAsian?: string;
  /**
   * Complex-script font family. Sets `<a:cs>`, the typeface renderers pick
   * for complex scripts such as Arabic, Hebrew and Thai — independently of
   * `font` (Latin) and `fontEastAsian` (CJK).
   */
  fontComplexScript?: string;
  /** Font size in points; fractional values allowed (`12`, `12.5`). */
  size?: number;
  /**
   * Color. Accepts `#RRGGBB`, `RRGGBB`, an ECMA-376 scheme color token
   * (`tx1`, `accent1`, ...), or `null` to clear.
   */
  color?: Color | null;
  /**
   * Ordered adjustments to `color` (`<a:lumMod>`, `<a:tint>`, `<a:alpha>`,
   * ...), written as children of its color element — the reference desktop app's theme
   * tints, such as Accent 2 Lighter 60%. Same field as on gradient stops;
   * requires `color`, and replaces any transforms the run's color had.
   */
  colorTransforms?: readonly ColorTransform[];
  /**
   * Non-solid glyph fill. The `color` field remains the shorthand for a solid
   * fill. Both are the same OOXML fill choice, so supplying both is rejected.
   */
  textFill?: TextFill;
  /**
   * `null` removes `b`, so the run inherits its weight again; `false` writes
   * `b="0"`, which overrides an inherited bold.
   */
  bold?: boolean | null;
  italic?: boolean;
  /**
   * Underline style. `true` is shorthand for `'sng'` (single). Pass the
   * exact `ST_TextUnderlineType` token for other styles (`'dbl'`, `'wavy'`,
   * `'dash'`, ...).
   */
  underline?: boolean | string;
  /**
   * Underline color. `null` explicitly follows the run's text color via
   * `<a:uFillTx>`; omit the property to leave an existing underline fill
   * untouched.
   */
  underlineColor?: Color | null;
  /**
   * Strikethrough style. `true` is shorthand for `'sngStrike'` (single
   * line). Pass the exact `ST_TextStrikeType` token (`'sngStrike'`,
   * `'dblStrike'`, `'noStrike'`) for other styles. `false` clears.
   */
  strike?: boolean | string;
  /**
   * Character spacing in 1/100 points (`0` = default). Negative values
   * tighten, positive values loosen. Mirrors `<a:rPr spc="…"/>`. `null`
   * removes it, so the run inherits its spacing again.
   */
  spc?: number | null;
  /**
   * Kerning threshold in 1/100 points (`ST_TextNonNegativePoint`, the same
   * unit as `spc`): `0` disables kerning, `1200` = apply kerning for runs
   * ≥12pt. Mirrors `<a:rPr kern="…"/>`. The reference desktop app on Mac was observed to
   * save the value as `kern="0"` when its Use kerning checkbox is cleared.
   */
  kern?: number;
  /** Normalize character heights (`<a:rPr normalizeH>`; the reference desktop app's Equalize character height). */
  normalizeHeight?: boolean;
  /**
   * Baseline offset as a fraction of 1 (`0.3` = superscript ~30% up,
   * `-0.25` = subscript). The reference desktop app emits ST_Percentage; this getter
   * returns the unit-fraction form for ergonomic comparisons.
   */
  baseline?: number;
  /**
   * Capitalization mode: `'none'`, `'small'` (smallCaps), or `'all'`
   * (allCaps). Mirrors `<a:rPr cap="…"/>`.
   */
  cap?: 'none' | 'small' | 'all';
  /**
   * Highlight color (cell-fill style background per run). Same color
   * format as `color`. Mirrors `<a:rPr><a:highlight>…</a:highlight></a:rPr>`.
   */
  highlight?: Color | null;
  /**
   * Outline drawn around the glyphs — `<a:rPr><a:ln>`, the character-level
   * twin of `setShapeStroke`. `null` removes it, which is not the same as
   * an outline of width 0: removing restores what the run inherits.
   */
  outline?: TextOutline | null;
  /**
   * Drop shadow behind the glyphs — `<a:outerShdw>` in the run's own
   * `<a:effectLst>`. `null` removes it.
   */
  shadow?: ShadowOptions | null;
  /** Inner shadow inside the glyphs — `<a:innerShdw>` in the run's effect list. */
  innerShadow?: InnerShadowOptions | null;
  /**
   * Glow around the glyphs — `<a:glow>` in the run's own `<a:effectLst>`.
   * `null` removes it.
   */
  glow?: GlowOptions | null;
  /** Reflection below the glyphs — `<a:reflection>` in the run's effect list. */
  reflection?: ReflectionOptions | null;
}
/**
 * Non-solid glyph fill. `none` is `<a:noFill>` (glyphs show only their
 * outline); `image` is `<a:blipFill>`, the picture stretched over the text
 * block — the reference desktop app's Picture or texture fill. Its `bytes` must be PNG, JPEG,
 * GIF, BMP, TIFF or WebP, and each edit embeds them once per part.
 */
type TextFill = ({
  readonly kind: 'gradient';
} & GradientFillOptions) | ({
  readonly kind: 'pattern';
} & PatternFillOptions) | {
  readonly kind: 'none';
} | {
  readonly kind: 'image';
  readonly bytes: Uint8Array;
};
type ReadTextFill = ({
  readonly kind: 'gradient';
} & ReadGradientFill) | {
  readonly kind: 'pattern';
  readonly preset: PatternFillOptions['preset'];
  readonly foreground: string;
  readonly background: string;
  readonly foregroundTransforms?: PatternFillOptions['foregroundTransforms'];
  readonly backgroundTransforms?: PatternFillOptions['backgroundTransforms'];
} | {
  readonly kind: 'none';
} | {
  readonly kind: 'image';
  /**
   * The picture, when the reader can reach the part that holds it — the
   * run's own slide. A picture inherited from a layout or master, or one
   * behind a broken relationship, reads without bytes.
   */
  readonly bytes?: Uint8Array;
};
/** A run's outline: `CT_LineProperties` as far as text uses it. */
interface TextOutline {
  /** Same accepted forms as `TextFormat.color`. */
  readonly color?: Color;
  /** Ordered adjustments to `color`, as on `TextFormat.colorTransforms`. Requires `color`. */
  readonly colorTransforms?: readonly ColorTransform[];
  /** Line width in EMU. The reference desktop app's thinnest visible text outline is 9525 (0.75pt). */
  readonly widthEmu?: number;
  /**
   * Gradient outline (`<a:gradFill>` in the run's `<a:ln>`), as on
   * `setShapeStroke`. Exclusive with `color`.
   */
  readonly fill?: LineFill;
}
/**
 * A run format read back from a deck. `color` and `highlight` widen to
 * `string`: when no theme is supplied, or a token is not in the scheme, the
 * readers surface the raw `<a:schemeClr val>` token as-is.
 */
type ReadTextFormat = Omit<TextFormat, 'bold' | 'spc' | 'color' | 'textFill' | 'underlineColor' | 'highlight' | 'outline' | 'shadow' | 'innerShadow' | 'glow' | 'reflection'> & {
  bold?: boolean;
  spc?: number;
  color?: string | null;
  underlineColor?: string | null;
  highlight?: string | null;
  outline?: ReadTextOutline | null;
  shadow?: (Omit<ShadowOptions, 'color'> & {
    readonly color?: string;
  }) | null;
  innerShadow?: (Omit<InnerShadowOptions, 'color'> & {
    readonly color?: string;
  }) | null;
  glow?: (Omit<GlowOptions, 'color'> & {
    readonly color: string;
  }) | null;
  reflection?: ReflectionOptions | null;
  textFill?: ReadTextFill;
};
/** A run outline read back from a deck. `color` widens for the same reason. */
type ReadTextOutline = Omit<TextOutline, 'color' | 'fill'> & {
  readonly color?: string;
  readonly fill?: ReadLineFill;
};
/**
 * Converts a format read back from a deck into one a writer accepts. The
 * readers widen every color to `string`, because a deck can hold a scheme
 * token that is not in its theme; this checks each one and drops the property
 * (or, for a glow, the whole effect) when the writer would reject it, so the
 * round trip never writes a color the schema has no room for.
 */
declare const toWritableTextFormat: (format: ReadTextFormat) => TextFormat;
//#endregion
//#region src/internal/drawingml/text-body-mutation.d.ts
/**
 * Bullet style descriptor for a paragraph.
 *
 *   - `'bullet'` — shortcut for `{ char: '•' }`.
 *   - `'number'` — shortcut for `{ autoNum: 'arabicPeriod' }` (1., 2., 3.).
 *   - `'none'` — emits `<a:buNone/>`, forcing the paragraph to be bullet-free
 *     even if the layout has a bullet default.
 *   - `{ char }` — custom bullet character (any single grapheme).
 *   - `{ autoNum }` — any ECMA-376 `ST_TextAutonumberScheme` token
 *     (`arabicPeriod`, `romanLcParenR`, `alphaUcPeriod`, ...).
 */
type BulletStyle = 'bullet' | 'number' | 'none' | {
  char: string;
} | {
  autoNum: string;
};
/**
 * Paragraph horizontal alignment per `ST_TextAlignType` (ECMA-376
 * §17.18.59). The library accepts plain-English names and translates to
 * the spec tokens; advanced callers can pass the raw token directly.
 */
type ParagraphAlignment = 'left' | 'center' | 'right' | 'justify' | 'distribute' | ParagraphAlignmentToken;
declare const ALIGNMENT_TOKENS: readonly ["l", "ctr", "r", "just", "dist", "justLow", "thaiDist"];
/**
 * The `ST_TextAlignType` tokens themselves — what a literal `algn` read
 * returns. Setters take the wider `ParagraphAlignment`; reads never produce a
 * plain-English name, because `justLow` and `thaiDist` have none and a read
 * must not lose them.
 */
type ParagraphAlignmentToken = (typeof ALIGNMENT_TOKENS)[number];
/** One run of a `ParagraphSpec`: text plus the format applied to its `<a:rPr>`. */
interface RunSpec {
  /**
   * Run text, written verbatim — unlike `setShapeText`, newlines are not
   * split into paragraphs (the reference desktop app shows a CR LF inside a run as a line
   * break, which some exporters rely on).
   */
  readonly text: string;
  readonly format?: TextFormat;
}
/** One paragraph for `setTextBodyParagraphs`: optional alignment plus its runs. */
interface ParagraphSpec {
  readonly align?: ParagraphAlignment;
  readonly runs: ReadonlyArray<RunSpec>;
  /**
   * Format of the paragraph-end mark (`<a:endParaRPr>`), the place to give a
   * paragraph with no runs a font size. Omit to write no `<a:endParaRPr>`.
   */
  readonly endFormat?: TextFormat;
}
//#endregion
//#region src/internal/drawingml/text-3d.d.ts
/** ECMA-376 `ST_PresetCameraType`; the reference desktop app's flat text-art presets use `'orthographicFront'`. */
type CameraPreset = (typeof CAMERA_PRESETS)[number];
/** ECMA-376 `ST_LightRigType` (`'threePt'`, `'soft'`, `'harsh'`, ...). */
type LightRigType = (typeof LIGHT_RIG_TYPES)[number];
/** ECMA-376 `ST_LightRigDirection`: where the light rig sits relative to the scene. */
type LightRigDirection = (typeof LIGHT_RIG_DIRECTIONS)[number];
/** ECMA-376 `ST_BevelPresetType` (`'circle'`, `'angle'`, `'softRound'`, ...). */
type BevelPreset = (typeof BEVEL_PRESETS)[number];
/** ECMA-376 `ST_PresetMaterialType` (`'matte'`, `'softEdge'`, `'metal'`, ...). */
type PresetMaterial = (typeof PRESET_MATERIALS)[number];
/** `CT_SphereCoords`: a rotation in degrees, each in `[0, 360)`. */
interface Rotation3D {
  readonly latitudeDeg: number;
  readonly longitudeDeg: number;
  readonly revolutionDeg: number;
}
/** `<a:lightRig>`. */
interface LightRig {
  readonly type: LightRigType;
  readonly direction: LightRigDirection;
  /** `<a:rot>`; omitted, the rig keeps its preset orientation. */
  readonly rotation?: Rotation3D;
}
/** `<a:scene3d>`: the camera and light the 3-D text is rendered with. */
interface Scene3D {
  readonly camera: CameraPreset;
  /**
   * The camera's `<a:rot>`, which overrides the preset's orientation —
   * the reference desktop app's X / Y / Z Rotation fields.
   */
  readonly cameraRotation?: Rotation3D;
  /** The camera's field of view (`fov`) in degrees, `[0, 180]`; the reference desktop app's Perspective field. */
  readonly fieldOfViewDeg?: number;
  readonly lightRig: LightRig;
}
/** `CT_Bevel`. Omitted fields take the schema defaults: `circle`, 76200 × 76200 EMU. */
interface Bevel {
  readonly preset?: BevelPreset;
  readonly widthEmu?: number;
  readonly heightEmu?: number;
}
/**
 * 3-D on a text body — what the reference desktop app writes for text-art bevels. `scene`
 * becomes `<a:scene3d>`; the other fields become `<a:sp3d>`, written when any
 * of them is set.
 */
interface Text3D {
  readonly scene?: Scene3D;
  /** Top bevel (`<a:bevelT>`). */
  readonly bevelTop?: Bevel;
  /** Bottom bevel (`<a:bevelB>`). */
  readonly bevelBottom?: Bevel;
  /** Extrusion depth (`extrusionH`) in EMU. */
  readonly extrusionHeightEmu?: number;
  /** Extrusion color (`<a:extrusionClr>`). Same accepted forms as `TextFormat.color`. */
  readonly extrusionColor?: Color;
  /** Ordered adjustments to `extrusionColor`. Requires `extrusionColor`. */
  readonly extrusionColorTransforms?: readonly ColorTransform[];
  /** Contour width (`contourW`) in EMU. */
  readonly contourWidthEmu?: number;
  /** Distance from the ground (`z`) in EMU; may be negative. */
  readonly distanceFromGroundEmu?: number;
  /** Surface material (`prstMaterial`). */
  readonly material?: PresetMaterial;
  /** Contour color (`<a:contourClr>`). Same accepted forms as `TextFormat.color`. */
  readonly contourColor?: Color;
  /** Ordered adjustments to `contourColor`. Requires `contourColor`. */
  readonly contourColorTransforms?: readonly ColorTransform[];
}
/**
 * A 3-D read back from a deck; `contourColor` and `extrusionColor` are the raw
 * `#RRGGBB` or scheme token.
 */
type ReadText3D = Omit<Text3D, 'contourColor' | 'extrusionColor'> & {
  readonly contourColor?: string;
  readonly extrusionColor?: string;
};
//#endregion
//#region src/internal/drawingml/custom-geometry.d.ts
/** A point in the path's own coordinate space, fully evaluated. */
interface GeomPoint {
  readonly x: number;
  readonly y: number;
}
/**
 * `ST_PathFillMode` (ECMA-376 §20.1.10.38). `none` means the path is
 * stroke-only; every other value fills (the lighten/darken variants only
 * affect how the fill is shaded, which the preview does not model, so we
 * still surface them rather than collapsing to `norm`). Default `norm`.
 */
type PathFillMode = 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
/**
 * One pen command. Discriminated on `kind` so renderers can `switch`
 * exhaustively. All coordinates / radii are in the owning path's
 * coordinate space; angles are in 60000ths of a degree (ECMA-376's unit
 * for `<a:arcTo>` `stAng` / `swAng`).
 */
type GeomCommand = {
  readonly kind: 'moveTo';
  readonly pt: GeomPoint;
} | {
  readonly kind: 'lnTo';
  readonly pt: GeomPoint;
} | {
  readonly kind: 'arcTo';
  readonly wR: number;
  readonly hR: number;
  readonly stAng: number;
  readonly swAng: number;
} | {
  readonly kind: 'quadBezTo';
  readonly pts: readonly [GeomPoint, GeomPoint];
} | {
  readonly kind: 'cubicBezTo';
  readonly pts: readonly [GeomPoint, GeomPoint, GeomPoint];
} | {
  readonly kind: 'close';
};
interface GeomPath {
  /** Path coordinate-space width (`<a:path w>`), or `null` when omitted. */
  readonly w: number | null;
  /** Path coordinate-space height (`<a:path h>`), or `null` when omitted. */
  readonly h: number | null;
  readonly fill: PathFillMode;
  readonly stroke: boolean;
  readonly commands: readonly GeomCommand[];
}
/**
 * `<a:rect>` (§20.1.9.22) — the rectangle a shape lays its text into,
 * in EMU from the shape's own top-left corner.
 *
 * Unlike the path commands above, this is written in the guide coordinate
 * space rather than a path's, so a value is either a literal EMU number or a
 * guide name; both are already resolved here. Custom shapes that omit it lay
 * text into the whole box, which is what `null` means.
 */
interface GeomTextRect {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}
interface CustomGeometry {
  readonly paths: readonly GeomPath[];
  readonly textRect: GeomTextRect | null;
}
//#endregion
//#region src/internal/drawingml/builtin-picture-styles.generated.d.ts
/** The reference desktop app's built-in picture styles, in the order of its Picture Styles gallery. */
declare const BUILTIN_PICTURE_STYLES: readonly ["Simple Frame, White", "Beveled Matte, White", "Metal Frame", "Drop Shadow Rectangle", "Reflected Rounded Rectangle", "Soft Edge Rectangle", "Double Frame, Black", "Thick Matte, Black", "Simple Frame, Black", "Beveled Oval, Black", "Compound Frame, Black", "Moderate Frame, Black", "Center Shadow Rectangle", "Rounded Diagonal Corner, White", "Snip Diagonal Corner, White", "Moderate Frame, White", "Rotated, White", "Perspective Shadow, White", "Relaxed Perspective, White", "Soft Edge Oval", "Bevel Rectangle", "Bevel Perspective", "Reflected Perspective Right", "Bevel Perspective Left, White", "Reflected Bevel, Black", "Reflected Bevel, White", "Metal Rounded Rectangle", "Metal Oval"];
//#endregion
//#region src/internal/drawingml/picture-styles.d.ts
/** English name of one of the reference desktop app's built-in picture styles (its gallery tooltip). */
type BuiltinPictureStyleName = (typeof BUILTIN_PICTURE_STYLES)[number];
//#endregion
//#region src/internal/drawingml/geometry.d.ts
interface Position {
  /** Left edge offset from the slide origin, in EMU. */
  readonly x: number;
  /** Top edge offset from the slide origin, in EMU. */
  readonly y: number;
}
interface Size {
  /** Width in EMU. */
  readonly w: number;
  /** Height in EMU. */
  readonly h: number;
}
//#endregion
//#region src/internal/presentationml/builtin-table-styles.generated.d.ts
/** The reference desktop app’s built-in table styles, in the order of its Table Styles gallery. */
declare const BUILTIN_TABLE_STYLES: readonly [{
  readonly id: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";
  readonly name: "No Style, No Grid";
  readonly category: "bestMatch";
}, {
  readonly id: "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}";
  readonly name: "Themed Style 1 - Accent 1";
  readonly category: "bestMatch";
}, {
  readonly id: "{284E427A-3D55-4303-BF80-6455036E1DE7}";
  readonly name: "Themed Style 1 - Accent 2";
  readonly category: "bestMatch";
}, {
  readonly id: "{69C7853C-536D-4A76-A0AE-DD22124D55A5}";
  readonly name: "Themed Style 1 - Accent 3";
  readonly category: "bestMatch";
}, {
  readonly id: "{775DCB02-9BB8-47FD-8907-85C794F793BA}";
  readonly name: "Themed Style 1 - Accent 4";
  readonly category: "bestMatch";
}, {
  readonly id: "{35758FB7-9AC5-4552-8A53-C91805E547FA}";
  readonly name: "Themed Style 1 - Accent 5";
  readonly category: "bestMatch";
}, {
  readonly id: "{08FB837D-C827-4EFA-A057-4D05807E0F7C}";
  readonly name: "Themed Style 1 - Accent 6";
  readonly category: "bestMatch";
}, {
  readonly id: "{5940675A-B579-460E-94D1-54222C63F5DA}";
  readonly name: "No Style, Table Grid";
  readonly category: "bestMatch";
}, {
  readonly id: "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}";
  readonly name: "Themed Style 2 - Accent 1";
  readonly category: "bestMatch";
}, {
  readonly id: "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}";
  readonly name: "Themed Style 2 - Accent 2";
  readonly category: "bestMatch";
}, {
  readonly id: "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}";
  readonly name: "Themed Style 2 - Accent 3";
  readonly category: "bestMatch";
}, {
  readonly id: "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}";
  readonly name: "Themed Style 2 - Accent 4";
  readonly category: "bestMatch";
}, {
  readonly id: "{327F97BB-C833-4FB7-BDE5-3F7075034690}";
  readonly name: "Themed Style 2 - Accent 5";
  readonly category: "bestMatch";
}, {
  readonly id: "{638B1855-1B75-4FBE-930C-398BA8C253C6}";
  readonly name: "Themed Style 2 - Accent 6";
  readonly category: "bestMatch";
}, {
  readonly id: "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}";
  readonly name: "Light Style 1";
  readonly category: "light";
}, {
  readonly id: "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
  readonly name: "Light Style 1 - Accent 1";
  readonly category: "light";
}, {
  readonly id: "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}";
  readonly name: "Light Style 1 - Accent 2";
  readonly category: "light";
}, {
  readonly id: "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}";
  readonly name: "Light Style 1 - Accent 3";
  readonly category: "light";
}, {
  readonly id: "{D27102A9-8310-4765-A935-A1911B00CA55}";
  readonly name: "Light Style 1 - Accent 4";
  readonly category: "light";
}, {
  readonly id: "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}";
  readonly name: "Light Style 1 - Accent 5";
  readonly category: "light";
}, {
  readonly id: "{68D230F3-CF80-4859-8CE7-A43EE81993B5}";
  readonly name: "Light Style 1 - Accent 6";
  readonly category: "light";
}, {
  readonly id: "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}";
  readonly name: "Light Style 2";
  readonly category: "light";
}, {
  readonly id: "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}";
  readonly name: "Light Style 2 - Accent 1";
  readonly category: "light";
}, {
  readonly id: "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}";
  readonly name: "Light Style 2 - Accent 2";
  readonly category: "light";
}, {
  readonly id: "{F2DE63D5-997A-4646-A377-4702673A728D}";
  readonly name: "Light Style 2 - Accent 3";
  readonly category: "light";
}, {
  readonly id: "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}";
  readonly name: "Light Style 2 - Accent 4";
  readonly category: "light";
}, {
  readonly id: "{5A111915-BE36-4E01-A7E5-04B1672EAD32}";
  readonly name: "Light Style 2 - Accent 5";
  readonly category: "light";
}, {
  readonly id: "{912C8C85-51F0-491E-9774-3900AFEF0FD7}";
  readonly name: "Light Style 2 - Accent 6";
  readonly category: "light";
}, {
  readonly id: "{616DA210-FB5B-4158-B5E0-FEB733F419BA}";
  readonly name: "Light Style 3";
  readonly category: "light";
}, {
  readonly id: "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}";
  readonly name: "Light Style 3 - Accent 1";
  readonly category: "light";
}, {
  readonly id: "{5DA37D80-6434-44D0-A028-1B22A696006F}";
  readonly name: "Light Style 3 - Accent 2";
  readonly category: "light";
}, {
  readonly id: "{8799B23B-EC83-4686-B30A-512413B5E67A}";
  readonly name: "Light Style 3 - Accent 3";
  readonly category: "light";
}, {
  readonly id: "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}";
  readonly name: "Light Style 3 - Accent 4";
  readonly category: "light";
}, {
  readonly id: "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}";
  readonly name: "Light Style 3 - Accent 5";
  readonly category: "light";
}, {
  readonly id: "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}";
  readonly name: "Light Style 3 - Accent 6";
  readonly category: "light";
}, {
  readonly id: "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}";
  readonly name: "Medium Style 1";
  readonly category: "medium";
}, {
  readonly id: "{B301B821-A1FF-4177-AEE7-76D212191A09}";
  readonly name: "Medium Style 1 - Accent 1";
  readonly category: "medium";
}, {
  readonly id: "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}";
  readonly name: "Medium Style 1 - Accent 2";
  readonly category: "medium";
}, {
  readonly id: "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}";
  readonly name: "Medium Style 1 - Accent 3";
  readonly category: "medium";
}, {
  readonly id: "{1E171933-4619-4E11-9A3F-F7608DF75F80}";
  readonly name: "Medium Style 1 - Accent 4";
  readonly category: "medium";
}, {
  readonly id: "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}";
  readonly name: "Medium Style 1 - Accent 5";
  readonly category: "medium";
}, {
  readonly id: "{10A1B5D5-9B99-4C35-A422-299274C87663}";
  readonly name: "Medium Style 1 - Accent 6";
  readonly category: "medium";
}, {
  readonly id: "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}";
  readonly name: "Medium Style 2";
  readonly category: "medium";
}, {
  readonly id: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
  readonly name: "Medium Style 2 - Accent 1";
  readonly category: "medium";
}, {
  readonly id: "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}";
  readonly name: "Medium Style 2 - Accent 2";
  readonly category: "medium";
}, {
  readonly id: "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}";
  readonly name: "Medium Style 2 - Accent 3";
  readonly category: "medium";
}, {
  readonly id: "{00A15C55-8517-42AA-B614-E9B94910E393}";
  readonly name: "Medium Style 2 - Accent 4";
  readonly category: "medium";
}, {
  readonly id: "{7DF18680-E054-41AD-8BC1-D1AEF772440D}";
  readonly name: "Medium Style 2 - Accent 5";
  readonly category: "medium";
}, {
  readonly id: "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}";
  readonly name: "Medium Style 2 - Accent 6";
  readonly category: "medium";
}, {
  readonly id: "{8EC20E35-A176-4012-BC5E-935CFFF8708E}";
  readonly name: "Medium Style 3";
  readonly category: "medium";
}, {
  readonly id: "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}";
  readonly name: "Medium Style 3 - Accent 1";
  readonly category: "medium";
}, {
  readonly id: "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}";
  readonly name: "Medium Style 3 - Accent 2";
  readonly category: "medium";
}, {
  readonly id: "{EB344D84-9AFB-497E-A393-DC336BA19D2E}";
  readonly name: "Medium Style 3 - Accent 3";
  readonly category: "medium";
}, {
  readonly id: "{EB9631B5-78F2-41C9-869B-9F39066F8104}";
  readonly name: "Medium Style 3 - Accent 4";
  readonly category: "medium";
}, {
  readonly id: "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}";
  readonly name: "Medium Style 3 - Accent 5";
  readonly category: "medium";
}, {
  readonly id: "{2A488322-F2BA-4B5B-9748-0D474271808F}";
  readonly name: "Medium Style 3 - Accent 6";
  readonly category: "medium";
}, {
  readonly id: "{D7AC3CCA-C797-4891-BE02-D94E43425B78}";
  readonly name: "Medium Style 4";
  readonly category: "medium";
}, {
  readonly id: "{69CF1AB2-1976-4502-BF36-3FF5EA218861}";
  readonly name: "Medium Style 4 - Accent 1";
  readonly category: "medium";
}, {
  readonly id: "{8A107856-5554-42FB-B03E-39F5DBC370BA}";
  readonly name: "Medium Style 4 - Accent 2";
  readonly category: "medium";
}, {
  readonly id: "{0505E3EF-67EA-436B-97B2-0124C06EBD24}";
  readonly name: "Medium Style 4 - Accent 3";
  readonly category: "medium";
}, {
  readonly id: "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}";
  readonly name: "Medium Style 4 - Accent 4";
  readonly category: "medium";
}, {
  readonly id: "{22838BEF-8BB2-4498-84A7-C5851F593DF1}";
  readonly name: "Medium Style 4 - Accent 5";
  readonly category: "medium";
}, {
  readonly id: "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}";
  readonly name: "Medium Style 4 - Accent 6";
  readonly category: "medium";
}, {
  readonly id: "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}";
  readonly name: "Dark Style 1";
  readonly category: "dark";
}, {
  readonly id: "{125E5076-3810-47DD-B79F-674D7AD40C01}";
  readonly name: "Dark Style 1 - Accent 1";
  readonly category: "dark";
}, {
  readonly id: "{37CE84F3-28C3-443E-9E96-99CF82512B78}";
  readonly name: "Dark Style 1 - Accent 2";
  readonly category: "dark";
}, {
  readonly id: "{D03447BB-5D67-496B-8E87-E561075AD55C}";
  readonly name: "Dark Style 1 - Accent 3";
  readonly category: "dark";
}, {
  readonly id: "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}";
  readonly name: "Dark Style 1 - Accent 4";
  readonly category: "dark";
}, {
  readonly id: "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}";
  readonly name: "Dark Style 1 - Accent 5";
  readonly category: "dark";
}, {
  readonly id: "{AF606853-7671-496A-8E4F-DF71F8EC918B}";
  readonly name: "Dark Style 1 - Accent 6";
  readonly category: "dark";
}, {
  readonly id: "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}";
  readonly name: "Dark Style 2";
  readonly category: "dark";
}, {
  readonly id: "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}";
  readonly name: "Dark Style 2 - Accent 1/Accent 2";
  readonly category: "dark";
}, {
  readonly id: "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}";
  readonly name: "Dark Style 2 - Accent 3/Accent 4";
  readonly category: "dark";
}, {
  readonly id: "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}";
  readonly name: "Dark Style 2 - Accent 5/Accent 6";
  readonly category: "dark";
}];
//#endregion
//#region src/internal/presentationml/table-styles.d.ts
/** English name the reference desktop app writes as `styleName` for a built-in table style. */
type BuiltinTableStyleName = (typeof BUILTIN_TABLE_STYLES)[number]['name'];
/** One entry of {@link BUILTIN_TABLE_STYLES}. */
type BuiltinTableStyle = (typeof BUILTIN_TABLE_STYLES)[number];
//#endregion
//#region src/internal/presentationml/default-masters.d.ts
type MasterPlaceholderType = 'title' | 'body' | 'dt' | 'ftr' | 'sldNum';
/** Insert Placeholder's menu, in the reference desktop app's order. */
type LayoutPlaceholderKind = 'content' | 'verticalContent' | 'text' | 'verticalText' | 'picture' | 'chart' | 'table' | 'smartArt' | 'media' | 'onlineImage';
type NotesMasterPlaceholderType = 'hdr' | 'dt' | 'sldImg' | 'body' | 'ftr' | 'sldNum';
type HandoutMasterPlaceholderType = 'hdr' | 'dt' | 'ftr' | 'sldNum';
//#endregion
//#region src/internal/presentationml/transition-builder.d.ts
type TransitionSpeed = 'slow' | 'med' | 'fast';
/**
 * Every transition effect element name in `CT_SlideTransition`'s choice
 * (ECMA-376 pml.xsd), emitted verbatim as `<p:{token}/>`.
 */
/**
 * Every effect element `setSlideTransition` writes, ECMA-376 and extensions
 * alike. The list is both the type's domain and the write-time validation
 * domain — keeping them one declaration stops them drifting apart.
 */
declare const TRANSITION_EFFECTS: readonly ["blinds", "checker", "circle", "dissolve", "comb", "cover", "cut", "diamond", "fade", "newsflash", "plus", "pull", "push", "random", "randomBar", "split", "strips", "wedge", "wheel", "wipe", "zoom", "conveyor", "doors", "ferris", "flash", "flip", "flythrough", "gallery", "glitter", "honeycomb", "pan", "prism", "reveal", "ripple", "shred", "switch", "vortex", "warp", "wheelReverse", "window", "prstTrans", "morph"];
/**
 * Transition effect token. Maps to the `<p:{token}/>` (or `<p14:…>`,
 * `<p15:prstTrans>`, `<p159:morph>`) child of `<p:transition>`, except
 * `'none'`: that is the library-level sentinel for "no effect element", which
 * the schema's choice has no member for.
 */
type TransitionEffect = 'none' | (typeof TRANSITION_EFFECTS)[number];
/**
 * The `prst` names [MS-PPTX] §2.4.3.8 defines for `p15:prstTrans`. The
 * attribute is an `xsd:string` naming "the internal resource to use", so a
 * name outside this list is one the reference desktop app has no transition for.
 */
declare const TRANSITION_PRESETS: readonly ["fallOver", "drape", "curtains", "wind", "prestige", "fracture", "crush", "peelOff", "pageCurlDouble", "pageCurlSingle", "airplane", "origami"];
/** A `p15:prstTrans` preset name. */
type TransitionPreset = (typeof TRANSITION_PRESETS)[number];
/** ST_TransitionMorphOption ([MS-PPTX] §2.6.4.1): what Morph matches between slides. */
type MorphOption = 'byObject' | 'byWord' | 'byChar';
interface TransitionOptions {
  effect: TransitionEffect;
  /**
   * Effect speed: 0.5 s (`fast`), 0.75 s (`med`) or 1 s (`slow`). Omitted, it
   * is the schema default `fast` (ECMA-376 pml.xsd, CT_SlideTransition).
   */
  speed?: TransitionSpeed;
  /**
   * Direction, valid only for effects that carry a `dir` attribute and only
   * within that effect's domain (validated on write). The side tokens name
   * the way the slide moves, so the reference desktop app's "From Right" is `l`.
   *   - `blinds`/`checker`/`comb`/`randomBar`/`doors`/`window`: `horz` | `vert`
   *   - `push`/`wipe`/`vortex`/`pan`/`glitter`/`prism`: `l` | `r` | `u` | `d`
   *   - `cover`/`pull`: the above plus `lu` | `ru` | `ld` | `rd`
   *   - `strips`: `lu` | `ru` | `ld` | `rd`
   *   - `ripple`: `center` | `lu` | `ru` | `ld` | `rd`
   *   - `switch`/`flip`/`ferris`/`gallery`/`conveyor`/`reveal`: `l` | `r`
   *   - `zoom`/`split`/`warp`/`flythrough`/`shred`: `in` | `out`
   * A mismatched effect/direction pair throws; on any other effect a stray
   * `direction` is ignored.
   */
  direction?: string;
  /** For `wheel` and `wheelReverse`: number of spokes (unsigned integer). Omitted means 4. */
  spokes?: number;
  /** For `split`: orientation token (`horz` / `vert`). */
  orientation?: 'horz' | 'vert';
  /** For `fade`, `cut` and `reveal`: pass `true` to go through black. */
  thruBlack?: boolean;
  /**
   * For `glitter`: `diamond` (the default) or `hexagon`. For `shred`: `strip`
   * (the default) or `rectangle` — the reference desktop app's "Strips" and "Particles".
   */
  pattern?: 'diamond' | 'hexagon' | 'strip' | 'rectangle';
  /**
   * For `prism`: the slide's content turns rather than the slide itself
   * (the reference desktop app's Rotate and Orbit, against Cube and Box).
   */
  isContent?: boolean;
  /** For `prism`: the turn faces inwards (the reference desktop app's Box and Orbit). */
  isInverted?: boolean;
  /** For `flythrough`: the slides bounce at the end. */
  hasBounce?: boolean;
  /** For `prstTrans`: which preset transition. Required for that effect. */
  preset?: TransitionPreset;
  /** For `prstTrans`: mirror the preset left to right ("Right" instead of "Left"). */
  invertX?: boolean;
  /** For `prstTrans`: mirror the preset top to bottom. */
  invertY?: boolean;
  /** For `morph`: what is matched between the slides. Defaults to `byObject`. */
  morphOption?: MorphOption;
  /** Whether clicking advances; default `true` (the reference desktop app's default). */
  advanceOnClick?: boolean;
  /**
   * Milliseconds to auto-advance after this slide. Omit for click-only
   * advance.
   */
  advanceAfterMs?: number;
  /**
   * Effect duration in milliseconds (`p14:dur`, [MS-PPTX]). ECMA-376
   * only has the three `speed` steps, so the transition is written as
   * the reference desktop app writes it:
   *   - `spd` is the fastest speed at least this long (≤ 500 ms `fast`,
   *     ≤ 750 ms `med`, otherwise `slow`) unless `speed` is given; a derived
   *     `fast` is left out, being the schema default.
   *   - A duration equal to its speed's own (500, 750 or 1000 ms) is that
   *     speed alone: no `p14:dur`, and it reads back as `speed`.
   *   - Any other duration is an `mc:AlternateContent` whose `p14` choice
   *     carries `p14:dur` and whose fallback carries only `spd`.
   */
  durationMs?: number;
}
/**
 * A transition read back from a deck. `effect` widens to `string` because a
 * file authored elsewhere can carry an effect element this library does not
 * model; writing one still requires a `TransitionEffect`. `preset` widens the
 * same way.
 */
type SlideTransition = Omit<TransitionOptions, 'effect' | 'preset'> & {
  effect: string;
  preset?: string;
};
//#endregion
//#region src/internal/presentationml/animation-emphasis-options.d.ts
/** Which way Spin turns — the reference desktop app's Effect Options "Direction". */
type AnimationSpinDirection = 'clockwise' | 'counterclockwise';
/**
 * Which way Grow/Shrink scales: both axes, or only the width (`'horizontal'`)
 * or only the height (`'vertical'`).
 */
type AnimationScaleDirection = 'both' | 'horizontal' | 'vertical';
/**
 * The colour a colour emphasis effect changes to: a theme slot or an sRGB
 * value, optionally with DrawingML colour transforms — the tints and shades
 * of a theme colour palette are `lumMod` / `lumOff` on the theme slot.
 */
type AnimationColor = Color | {
  readonly color: Color;
  readonly colorTransforms?: readonly ColorTransform[];
};
//#endregion
//#region src/internal/presentationml/animation-builder.d.ts
/**
 * What kind of effect to apply — one token per entry of the reference desktop app's Entrance,
 * Emphasis and Exit galleries, named after the gallery entry.
 *
 * Entrance effects put the shape on the slide and end in `In` (`appear` aside);
 * exit effects take it off and end in `Out` (`disappear` aside). Emphasis
 * effects animate a shape that is already on the slide and never decide
 * whether it is shown.
 *
 * Where a gallery name would not tell two presets apart, the token does:
 * `zoomIn` is the reference desktop app's "Zoom" (preset 53, which fades as it grows) and
 * `basicZoomIn` its "Basic Zoom" (preset 23); `floatIn` / `floatOut` are
 * "Float In" / "Float Out" (preset 42) and `floatingIn` / `floatingOut` the
 * "Float" entrance and exit (preset 30).
 */
type AnimationEffect = 'appear' | 'blindsIn' | 'checkerboardIn' | 'dissolveIn' | 'flyIn' | 'peekIn' | 'randomBarsIn' | 'shapeIn' | 'splitIn' | 'stripsIn' | 'wedgeIn' | 'wheelIn' | 'wipeIn' | 'expandIn' | 'fadeIn' | 'swivelIn' | 'zoomIn' | 'centerRevolveIn' | 'floatIn' | 'growTurnIn' | 'riseUpIn' | 'spinnerIn' | 'basicZoomIn' | 'stretchIn' | 'boomerangIn' | 'bounceIn' | 'creditsIn' | 'curveUpIn' | 'dropIn' | 'flipIn' | 'floatingIn' | 'pinwheelIn' | 'spiralIn' | 'basicSwivelIn' | 'whipIn' | 'blindsOut' | 'checkerboardOut' | 'disappear' | 'dissolveOut' | 'flyOut' | 'peekOut' | 'randomBarsOut' | 'shapeOut' | 'splitOut' | 'stripsOut' | 'wedgeOut' | 'wheelOut' | 'wipeOut' | 'contractOut' | 'fadeOut' | 'swivelOut' | 'zoomOut' | 'centerRevolveOut' | 'collapseOut' | 'floatOut' | 'shrinkTurnOut' | 'sinkDownOut' | 'spinnerOut' | 'basicZoomOut' | 'stretchyOut' | 'boomerangOut' | 'bounceOut' | 'creditsOut' | 'curveDownOut' | 'dropOut' | 'flipOut' | 'floatingOut' | 'pinwheelOut' | 'spiralOut' | 'basicSwivelOut' | 'whipOut' | 'fillColor' | 'fontColor' | 'growShrink' | 'lineColor' | 'spin' | 'transparency' | 'boldFlash' | 'brushColor' | 'complementaryColor' | 'complementaryColor2' | 'contrastingColor' | 'darken' | 'desaturate' | 'lighten' | 'objectColor' | 'pulse' | 'underline' | 'colorPulse' | 'growWithColor' | 'shimmer' | 'teeter' | 'blink' | 'boldReveal' | 'wave';
/**
 * Which edge of the slide a `flyIn` comes from, or a `flyOut` leaves by —
 * the reference desktop app's "From Bottom" / Google Slides' "Fly in from bottom" — or, for
 * the four corners, the corner. A wipe or a peek names an edge too; strips
 * name a corner.
 */
type AnimationDirection = 'top' | 'right' | 'bottom' | 'left' | 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';
/** Which way blinds, random bars and a split run — and a checkerboard: across or down. */
type AnimationOrientation = 'horizontal' | 'vertical';
/** Whether a shape or split effect opens from the centre outwards or closes inwards. */
type AnimationInOut = 'in' | 'out';
/** The outline a `shapeIn` / `shapeOut` reveals the object through. */
type AnimationShape = 'circle' | 'box' | 'diamond' | 'plus';
/**
 * How an effect on a shape with text treats the text — the reference desktop app's Effect
 * Options "Sequence":
 *
 *   - `'asOneObject'` animates the shape and its text together, as one effect.
 *   - `'allAtOnce'` gives each paragraph an effect of its own, all starting
 *     together (`<p:bldP build="allAtOnce">`).
 *   - `'byParagraph'` gives each paragraph an effect of its own, one click (or
 *     start condition) apiece (`<p:bldP build="p">`).
 */
type AnimationTextBuild = 'asOneObject' | 'allAtOnce' | 'byParagraph';
/** When an effect runs relative to the one before it. */
type AnimationStartCondition = 'click' | 'withPrevious' | 'afterPrevious';
interface AnimationOptions {
  /** Which preset effect to apply. */
  readonly effect: AnimationEffect;
  /**
   * Which edge of the slide a `'flyIn'` comes from, or a `'flyOut'` leaves by
   * — any of the eight. `'wipeIn'`, `'wipeOut'`, `'peekIn'` and `'peekOut'`
   * take one of the four edges, `'stripsIn'` and `'stripsOut'` one of the four
   * corners. Defaults to the reference desktop app's own default for the preset (`'bottom'`;
   * `'bottomLeft'` for strips). Passing it for an effect that takes no
   * direction is an error rather than a no-op: it would otherwise read as a
   * direction the file never records. The same holds for the options below.
   */
  readonly direction?: AnimationDirection;
  /**
   * For blinds, random bars and split: which way the bars run. For a
   * checkerboard, `'horizontal'` is the reference desktop app's "Across" and `'vertical'` its
   * "Down". Defaults to `'horizontal'` (`'vertical'` for split).
   */
  readonly orientation?: AnimationOrientation;
  /**
   * For shape and split: open from the centre (`'out'`) or close in on it
   * (`'in'`). Defaults to `'in'` — `'out'` for `'shapeOut'`, as in the reference desktop app.
   */
  readonly inOut?: AnimationInOut;
  /** For `'shapeIn'` / `'shapeOut'`: the outline. Defaults to `'circle'`. */
  readonly shape?: AnimationShape;
  /** For `'wheelIn'` / `'wheelOut'`: 1, 2, 3, 4 or 8 spokes. Defaults to 1. */
  readonly spokes?: number;
  /** For `'spin'`: which way it turns. Defaults to `'clockwise'`. */
  readonly spinDirection?: AnimationSpinDirection;
  /**
   * For `'spin'`: how far it turns, in degrees. The reference desktop app's menu offers 90
   * (Quarter Spin), 180 (Half Spin), 360 (Full Spin, the default) and 720
   * (Two Spins); any positive angle can be written.
   */
  readonly spinDegrees?: number;
  /**
   * For `'growShrink'`: scale both axes (the default), or only the width
   * (`'horizontal'`) or only the height (`'vertical'`).
   */
  readonly scaleDirection?: AnimationScaleDirection;
  /**
   * For `'growShrink'`: the size to reach, in percent of the shape's own.
   * The reference desktop app's menu offers 25 (Tiny), 50 (Smaller), 150 (Larger, the
   * default) and 400 (Huge).
   */
  readonly scalePercent?: number;
  /**
   * For `'transparency'`: how transparent the shape becomes, 0–100. The reference desktop app
   * offers 25, 50 (the default), 75 and 100.
   */
  readonly transparencyPercent?: number;
  /**
   * For the colour emphasis effects — `'fillColor'`, `'fontColor'`,
   * `'lineColor'`, `'brushColor'`, `'objectColor'`, `'colorPulse'` and
   * `'growWithColor'` — the colour to change to. Defaults to Accent 2, and to
   * Background 1 for `'colorPulse'`, as the reference desktop app's gallery does.
   */
  readonly color?: AnimationColor;
  /**
   * How long the effect runs, in milliseconds. Defaults to the reference desktop app's default
   * for the preset (`defaultAnimationDurationMs`). An effect made of several
   * behaviours is scaled as a whole, the way the reference desktop app's Duration box does.
   * `'appear'` and `'disappear'` are instantaneous and write no timed
   * behaviour, so it does not reach them; `'transparency'` and `'boldReveal'`
   * hold until the end of the slide and take none.
   */
  readonly durationMs?: number;
  /**
   * What starts the effect. Defaults to `'click'` — the effect waits for the
   * viewer's next click. `'withPrevious'` runs it alongside the effect before
   * it, `'afterPrevious'` once that effect has finished. When the effect is
   * the slide's first, neither has a predecessor to follow, so both start as
   * soon as the slide appears instead of waiting for a click.
   */
  readonly start?: AnimationStartCondition;
  /** How long to wait once the start condition is met. Defaults to 0ms. */
  readonly delayMs?: number;
  /**
   * How the shape's text is animated — the reference desktop app's Effect Options "Sequence".
   * Defaults to `'asOneObject'`, the gallery's own default. `'allAtOnce'` and
   * `'byParagraph'` give every paragraph an effect of its own: started together
   * for the first, each on its own start (a click apiece by default) for the
   * second. Fill Color and Line Color animate the shape and take neither.
   */
  readonly build?: AnimationTextBuild;
}
//#endregion
//#region src/internal/presentationml/comments-part.d.ts
interface CommentAuthor {
  /** Unique within the author list (also referenced by `<p:cm authorId="…">`). */
  readonly id: number;
  readonly name: string;
  readonly initials: string;
  /** Highest `idx` value this author has used; bumped on every new comment. */
  readonly lastIdx: number;
  /** Optional color index into the package's theme; preserved verbatim. */
  readonly clrIdx: number | null;
}
interface CommentPosition {
  readonly x: number;
  readonly y: number;
}
interface SlideComment {
  readonly parent?: {
    readonly authorId: number;
    readonly idx: number;
  };
  /** Author id (matches `CommentAuthor.id`). */
  readonly authorId: number;
  /** Per-author monotonic index. */
  readonly idx: number;
  /** ISO-8601 timestamp string as it appeared on disk; may be `null`. */
  readonly dt: string | null;
  readonly text: string;
  readonly position: CommentPosition | null;
}
//#endregion
//#region src/internal/presentationml/modern-comments-part.d.ts
/** `ST_CommentStatus` — the three states a thread can be in. */
type CommentStatus = 'active' | 'resolved' | 'closed';
declare const COMMENT_STATUSES: readonly CommentStatus[];
//#endregion
//#region src/api/_internal-symbols.d.ts
declare const INTERNAL_PACKAGE: unique symbol;
declare const SLIDE_PART_NAME: unique symbol;
declare const SLIDE_DOCUMENT: unique symbol;
declare const SLIDE_PART: unique symbol;
declare const SLIDE_SHAPES: unique symbol;
declare const SHAPE_SLIDE: unique symbol;
declare const SHAPE_ELEMENT: unique symbol;
declare const SHAPE_SNAPSHOT: unique symbol;
declare const LAYOUT_PART_NAME: unique symbol;
declare const LAYOUT_PART: unique symbol;
declare const LAYOUT_DOCUMENT: unique symbol;
declare const COMMENT_PARENT: unique symbol;
declare const COMMENT_SLIDE: unique symbol;
declare const COMMENT_SNAPSHOT: unique symbol;
declare const COMMENT_MODERN: unique symbol;
declare const CELL_TABLE: unique symbol;
declare const CELL_ELEMENT: unique symbol;
declare const CELL_ROW: unique symbol;
declare const CELL_COL: unique symbol;
/**
 * Opaque handle to a loaded / created presentation. Constructed via
 * `loadPresentation` / `createPresentation`; passed positionally to
 * every fn-API helper.
 */
interface PresentationData {
  readonly [INTERNAL_PACKAGE]: OpcPackage;
  /** @internal — populated lazily by `getSlides`. */
  _slidesCache: unknown[] | null;
}
/** Data shape backing every `Slide` value. */
interface SlideData {
  readonly [INTERNAL_PACKAGE]: OpcPackage;
  readonly [SLIDE_PART_NAME]: PartName;
  [SLIDE_DOCUMENT]: XmlDocument;
  [SLIDE_PART]: SlidePart;
  [SLIDE_SHAPES]: SlideShapeData[];
}
/** Data shape backing every `SlideShape` value. */
interface SlideShapeData {
  readonly [SHAPE_SLIDE]: SlideData;
  [SHAPE_ELEMENT]: XmlElement;
  [SHAPE_SNAPSHOT]: SlidePart['shapes'][number];
}
/**
 * Data shape backing every `SlideLayout` value. Carries the parsed
 * document, not just the typed view, so a layout handle can be edited
 * and written back the way a slide handle can.
 */
interface SlideLayoutData {
  readonly [INTERNAL_PACKAGE]: OpcPackage;
  readonly [LAYOUT_PART_NAME]: PartName;
  [LAYOUT_DOCUMENT]: XmlDocument;
  [LAYOUT_PART]: SlideLayoutPart;
}
/** Where a modern comment handle points: its thread, and itself within it. */
interface ModernCommentRef {
  /** Id of the `<p188:cm>`; equal to `id` when the handle is the thread itself. */
  readonly threadId: string;
  readonly id: string;
  readonly status: CommentStatus | null;
  readonly authorId: string;
}
/**
 * Opaque handle for one comment on a slide. The `author` is resolved
 * on read: against the package-level `commentAuthors.xml` for an
 * ECMA-376 `<p:cm>`, and against `authors.xml` for a modern one.
 */
interface SlideCommentData {
  [COMMENT_PARENT]: SlideCommentData | null;
  readonly [COMMENT_SLIDE]: SlideData;
  /**
   * For a modern comment this is a projection, not what is on disk: `text`,
   * `dt` (the thread's `created`) and `position` are real, while `authorId`
   * and `idx` are the author's place in the modern author list and the
   * comment's place in its part. They exist so the deck-wide queries read
   * one shape whatever the file uses; nothing is ever written back through
   * them. `COMMENT_MODERN` is what every modern mutation goes through.
   */
  [COMMENT_SNAPSHOT]: SlideComment;
  /** Set for a modern comment; `null` for an ECMA-376 one. */
  [COMMENT_MODERN]: ModernCommentRef | null;
  readonly author: CommentAuthor;
}
/**
 * Opaque handle for one cell of a table graphic-frame shape. Carries
 * its zero-based row/column position plus the parent table shape so
 * mutations can commit back through the standard slide path.
 */
interface TableCellData {
  readonly [CELL_TABLE]: SlideShapeData;
  readonly [CELL_ELEMENT]: XmlElement;
  readonly [CELL_ROW]: number;
  readonly [CELL_COL]: number;
}
//#endregion
//#region src/api/fn/package-io.d.ts
/**
 * Anything that can be turned into a `Uint8Array` of PPTX bytes:
 *
 *   - `Uint8Array` — used as-is.
 *   - `ArrayBuffer` — wrapped without copying.
 *   - `Blob` or `File` — read via `arrayBuffer()`.
 */
type PresentationInput = Uint8Array | ArrayBuffer | Blob;
/**
 * Loads an existing `.pptx` and returns a `PresentationData` value.
 */
declare const loadPresentation: (input: PresentationInput) => Promise<PresentationData>;
/** Slide-canvas aspect ratio for {@link createPresentation}. */
type PresentationSize = '16:9' | '4:3';
/**
 * Creates a fresh, immediately-authorable `PresentationData`.
 *
 * The returned deck carries a slide master, the default theme, and three
 * slide layouts — `'Blank'`, `'Title Slide'`, and `'Title and Content'` —
 * but no slides yet. Add one with {@link addSlide} (or the
 * `addBlankSlide` / `addTitleSlide` / `addContentSlide` helpers), then
 * `savePresentation` to emit a `.pptx` that opens cleanly in the reference desktop app,
 * Keynote, Google Slides, and LibreOffice.
 *
 * @param options.size Slide aspect ratio. `'16:9'` (12192000×6858000 EMU,
 *   The reference desktop app's modern default) or `'4:3'` (9144000×6858000). Defaults to
 *   `'16:9'`.
 */
declare const createPresentation: (options?: {
  size?: PresentationSize;
}) => PresentationData;
/**
 * Serializes a presentation back to PPTX bytes.
 */
declare const savePresentation: (pres: PresentationData) => Promise<Uint8Array>;
//#endregion
//#region src/api/fn/sections.d.ts
/** One section in the deck. `slides` is a snapshot at read time. */
interface SlideSection {
  readonly name: string;
  readonly slides: ReadonlyArray<SlideData>;
}
/**
 * Returns every section defined in `<p14:sectionLst>`, with each
 * section's slides resolved to `SlideData` handles via the
 * presentation's `<p:sldIdLst>`. Sections referencing missing slides
 * are dropped silently; empty arrays are preserved.
 *
 * Returns an empty array when no sectionLst is present.
 */
declare const getSlideSections: (pres: PresentationData) => ReadonlyArray<SlideSection>;
/**
 * Replaces the deck's section list with `sections`. Each section is
 * given a fresh GUID `id` attribute (the reference desktop app generates one per
 * section; we synthesize a deterministic-ish one based on index +
 * timestamp for v1).
 *
 * Pass `[]` to clear all sections — the helper drops the
 * `<p14:sectionLst>` extension entirely when no sections remain.
 */
declare const setSlideSections: (pres: PresentationData, sections: ReadonlyArray<{
  name: string;
  slides: ReadonlyArray<SlideData>;
}>) => void;
//#endregion
//#region src/api/fn/shape-read-base.d.ts
declare const getShapeKind: (shape: SlideShapeData) => ShapeKind;
declare const getShapeId: (shape: SlideShapeData) => number;
/**
 * Returns the preset-geometry token (`'rect'`, `'ellipse'`, `'star5'`,
 * `'rightArrow'`, ...) for shapes whose body carries a
 * `<a:prstGeom prst="…"/>`. Returns `null` for:
 *
 *   - group shapes, tables, and charts (they have no preset geometry),
 *   - shapes using custom geometry (`<a:custGeom>`),
 *   - shapes whose preset is missing (malformed but possible).
 *
 * Useful for renderers / inspectors that want to draw a faithful
 * approximation of each shape without dropping to the raw XML.
 */
declare const getShapePreset: (shape: SlideShapeData) => string | null;
/**
 * Reads the preset's adjust-handle values (`<a:prstGeom><a:avLst>
 * <a:gd name="adj" fmla="val 30000"/></a:avLst>`) as a map from guide
 * name → numeric value. Per ECMA-376 §20.1.9.4, guides are stored
 * with a formula prefix — `val 12345` is a literal number, and the
 * other prefixes (`pin`, `+-`, etc.) compute from other guides. We
 * only surface the `val` form because other formulas reference the
 * preset's built-in guides and don't make sense without them.
 *
 * Returns an empty record when no adjust values are authored (the
 * shape paints at its preset defaults).
 */
declare const getShapeAdjustValues: (shape: SlideShapeData) => Record<string, number>;
/**
 * Reads a shape's custom geometry (`<a:custGeom>`, ECMA-376 §20.1.9) as a
 * fully-evaluated path list. Returns `null` when the shape has no
 * `<a:custGeom>` in its own `<p:spPr>` (it uses a preset, inherits
 * geometry, or isn't a drawable shape), and also `null` when the geometry
 * is present but malformed — an unresolved guide reference, an unknown
 * `fmla` operator, or a command missing its points — so callers can fall
 * back rather than emit broken output.
 *
 * Guide formulas in `<a:avLst>` / `<a:gdLst>` are evaluated in document
 * order against the implicit built-in guides (`w`, `h`, `ss`, `hc`, the
 * `cdN` angle constants, …) derived from the shape extents. Per spec,
 * division by zero yields `0`. The returned commands carry only numbers —
 * no guide names survive.
 *
 * Pairs with `getShapePreset` (which returns `null` for custom-geometry
 * shapes): a renderer calls `getShapePreset` first, and falls through to
 * this when it gets `null`.
 */
declare const getShapeCustomGeometry: (shape: SlideShapeData) => CustomGeometry | null;
/**
 * Evaluates a preset shape's outline (`<a:prstGeom>`, ECMA-376 §20.1.9.18)
 * at `size`, from the standard's own preset definitions: the same paths,
 * shading modes (`lighten` / `darken` …) and text rectangle the reference desktop app uses.
 * The result has the shape of {@link getShapeCustomGeometry}'s, so a
 * renderer draws presets and custom geometry the same way.
 *
 * `adjustValues` are the shape's adjust handles (`getShapeAdjustValues`);
 * each replaces the preset default of the same name. Returns `null` for a
 * preset name ECMA-376 does not define.
 */
declare const getPresetGeometry: (preset: PresetShape | string, size: {
  readonly w: Emu;
  readonly h: Emu;
}, adjustValues?: Readonly<Record<string, number>>) => CustomGeometry | null;
/**
 * Returns the highest `cNvPr@id` used by any shape on the slide,
 * or `0` when the slide has no shapes with non-negative ids.
 *
 * Useful when hand-rolling a custom shape and you need an id known
 * not to collide. The official allocator inside `addSlideShape` /
 * `addSlideTextBox` etc. already does this — call those instead
 * when you don't need a custom id.
 */
declare const getMaxShapeId: (slide: SlideData) => number;
/**
 * Deck-wide sibling of `getMaxShapeId`. Returns the highest
 * `cNvPr@id` across every shape on every slide, or `0` when the
 * deck has no shapes.
 *
 * Note: shape ids are scoped to a slide in OOXML — collisions
 * across slides are fine. This helper is for the rare cases where
 * a caller wants a single id known to be higher than anything in
 * the deck (e.g. to keep ids monotonically increasing).
 */
declare const getMaxShapeIdInPresentation: (pres: PresentationData) => number;
/**
 * Returns the number of slide masters declared in the
 * presentation's `<p:sldMasterIdLst>`. Most decks use exactly one
 * master; multi-master decks come from templates that combine
 * brand variants (e.g. a corporate master + a sponsor master).
 *
 * Returns `0` if `presentation.xml` is missing.
 */
declare const getSlideMasterCount: (pres: PresentationData) => number;
/**
 * Returns the package part name of every slide master declared in
 * `presentation.xml`, resolved through the presentation's `.rels`.
 * Sibling of `getSlideMasterCount` for downstream tooling that
 * needs the master URIs (e.g. byte-level diff, custom validators).
 *
 * Returns an empty array when `presentation.xml` or its `.rels`
 * are missing.
 */
/**
 * Returns the part name of the slide master a slide or layout inherits from
 * (`/ppt/slideMasters/slideMaster1.xml`), or `null` when it has no master rel.
 *
 * Useful for multi-master decks where different slides live under
 * different brand templates and the caller needs to scope theme /
 * fontScheme / clrMap lookups to the correct master.
 */
declare const getSlideMasterPartName: (value: SlideData | SlideLayoutData) => string | null;
declare const getSlideMasterPartNames: (pres: PresentationData) => ReadonlyArray<string>;
/**
 * Histogram of master part name → number of slides whose layout chains
 * up to that master. Every master in the package appears as a key
 * (count `0` for unreferenced masters), so the function surfaces
 * unused masters directly — useful for trimming multi-master template
 * decks. Pair with `getSlideLayoutUsageCounts` for the layout layer.
 */
declare const getSlideMasterUsageCounts: (pres: PresentationData) => Readonly<Record<string, number>>;
/**
 * Master part names that no slide chains to (count of `0` in
 * `getSlideMasterUsageCounts`). Pair to `getUnusedSlideLayouts`.
 * Useful when trimming multi-master template decks of dead theme
 * variants.
 */
declare const getUnusedSlideMasters: (pres: PresentationData) => ReadonlyArray<string>;
declare const getShapeName: (shape: SlideShapeData) => string;
/**
 * Renames the shape's `cNvPr@name`. The display name is what
 * the reference desktop app shows in the Selection Pane and what `findShapeByName`
 * matches on. Empty strings are allowed (matches the reference desktop app's behavior).
 */
/**
 * Reads the shape's alt-text description (`<p:cNvPr descr="...">`).
 * Accessibility tools (screen readers, contrast checkers) and
 * the reference desktop app's "Alt Text" pane look at this field. Returns `null`
 * when no description is set.
 */
declare const getShapeDescription: (shape: SlideShapeData) => string | null;
/**
 * Sets the shape's alt-text description (`<p:cNvPr descr="...">`).
 * Pass `null` to clear. Important for accessibility — image shapes
 * and decorative graphics should carry a descr that conveys the
 * visual meaning to screen readers.
 */
declare const setShapeDescription: (shape: SlideShapeData, description: string | null) => void;
/**
 * Reads the shape's alt-text title (`<p:cNvPr title="...">`).
 * The reference desktop app surfaces this alongside `descr` in its Alt Text pane
 * as a short heading. Returns `null` when no title is set.
 */
declare const getShapeAltTitle: (shape: SlideShapeData) => string | null;
/**
 * Sets the shape's alt-text title (`<p:cNvPr title="...">`). Pass
 * `null` to clear. Distinct from `renameShape`, which writes the
 * `name` attribute used in the selection pane.
 */
declare const setShapeAltTitle: (shape: SlideShapeData, title: string | null) => void;
/**
 * `true` when the shape's `<p:cNvPr hidden="1">` is set. Hidden
 * shapes are skipped by the reference desktop app's renderer but stay in the
 * shape tree — useful for variant slides that toggle which boxes
 * are visible.
 */
declare const isShapeHidden: (shape: SlideShapeData) => boolean;
/**
 * Sets or clears `<p:cNvPr hidden="...">` on the shape. Hidden
 * shapes remain in the document but the reference desktop app doesn't render them.
 */
declare function setShapeHidden(shape: SlideShapeData, hidden: boolean): void;
declare function setShapeHidden(shapes: readonly SlideShapeData[], hidden: boolean): void;
/** Renames the shape: the name the Selection Pane shows (`<p:cNvPr name>`). */
declare const renameShape: (shape: SlideShapeData, newName: string) => void;
declare const getShapePlaceholderType: (shape: SlideShapeData) => string | null;
declare const getShapePlaceholderIdx: (shape: SlideShapeData) => number | null;
/**
 * `true` when the shape carries `<p:nvSpPr><p:nvPr><p:ph>` — i.e. it
 * inherits from a layout/master placeholder. False for decorative
 * geometry the slide author dropped onto the canvas. Decoupled from
 * the more specific `getShapePlaceholderType` / `getShapePlaceholderIdx`
 * (either can be null on a real placeholder; together they identify it).
 */
declare const isShapePlaceholder: (shape: SlideShapeData) => boolean;
/**
 * `true` when the shape is a **text box** (`<p:cNvSpPr txBox="1">`) rather than
 * an autoshape. The distinction drives default text formatting: the reference desktop app and
 * LibreOffice left-align and top-anchor text-box text, but center-align and
 * middle-anchor autoshape text when the deck author left those unset.
 */
declare const isShapeTextBox: (shape: SlideShapeData) => boolean;
declare const getShapeText: (shape: SlideShapeData) => string;
declare const getShapePosition: (shape: SlideShapeData) => Position | null;
declare const getShapeSize: (shape: SlideShapeData) => Size | null;
declare const getShapeRotation: (shape: SlideShapeData) => number;
declare const getShapeFlip: (shape: SlideShapeData) => {
  horizontal: boolean;
  vertical: boolean;
} | null;
/**
 * Enumerates the shapes nested inside a `<p:grpSp>` group, one level
 * deep (nested groups come through as `kind: 'group'` themselves —
 * call this again on each one to recurse).
 *
 * Returns an empty array for non-group shapes. Each child carries its
 * own bounds in the group's *internal* coordinate system; pair with
 * `getGroupTransform` to project them onto the slide.
 */
declare const getGroupChildren: (shape: SlideShapeData) => ReadonlyArray<SlideShapeData>;
/**
 * Returns the group's slide-relative bounds (`outer`) and the internal
 * coordinate system the children's `<a:xfrm>` values live in
 * (`inner`). Renderers project a child point `(cx, cy)` onto the slide
 * with:
 *
 *   slideX = outer.x + (cx - inner.x) * (outer.w / inner.w)
 *   slideY = outer.y + (cy - inner.y) * (outer.h / inner.h)
 *
 * Returns `null` for non-group shapes or for groups whose
 * `<p:grpSpPr>` omits an `<a:xfrm>`.
 */
declare const getGroupTransform: (shape: SlideShapeData) => {
  readonly outer: ShapeBounds;
  readonly inner: ShapeBounds;
} | null;
/**
 * Combined bounds — position + size in one object. Returns `null` when
 * the shape inherits both position and size from its layout (so the
 * `<a:xfrm>` element is absent or incomplete).
 */
interface ShapeBounds {
  readonly x: Emu;
  readonly y: Emu;
  readonly w: Emu;
  readonly h: Emu;
}
declare const getShapeBounds: (shape: SlideShapeData) => ShapeBounds | null;
/**
 * Same as `getShapeBounds` but walks the placeholder inheritance chain
 * when the shape has no `<a:xfrm>` of its own:
 *
 *   1. The shape's own bounds.
 *   2. The matching placeholder on the slide's layout (by `<p:ph idx>`,
 *      falling back to `<p:ph type>`).
 *   3. The matching placeholder on the layout's slide master.
 *
 * This is what renderers want when they need to draw a placeholder
 * that the deck author left unsized — real templates only override
 * geometry per placeholder when it differs from the master.
 *
 * Returns `null` when none of the three levels carries explicit bounds.
 */
declare const getShapeBoundsResolved: (pres: PresentationData, shape: SlideShapeData) => ShapeBounds | null;
/**
 * Returns the center point of the shape's bounds in EMU, or `null`
 * when the shape has no `<a:xfrm>`. Convenience for layout
 * pipelines that compute alignment / overlap from center points.
 */
declare const getShapeCenter: (shape: SlideShapeData) => {
  readonly x: Emu;
  readonly y: Emu;
} | null;
/**
 * `true` when point `(x, y)` (in EMU) lies inside the shape's
 * axis-aligned bounds. Closed on the top-left edge, open on the
 * bottom-right (standard half-open rectangle). Returns `false`
 * when the shape has no bounds.
 *
 * Useful for hit-testing in custom interaction handlers.
 */
declare const pointInShape: (shape: SlideShapeData, x: number, y: number) => boolean;
/**
 * Returns every shape on the slide whose bounds contain `(x, y)`
 * (in EMU). Built on `pointInShape`. The list is in document
 * order, so callers can index by z-stack from front (last) to
 * back (first) if they want one-hit semantics.
 */
declare const findShapesAtPoint: (slide: SlideData, x: number, y: number) => ReadonlyArray<SlideShapeData>;
/**
 * Returns every shape on the slide whose bounds overlap the rectangle
 * `(x, y, w, h)` (in EMU). "Overlap" means the rectangles intersect at
 * all — touching edges count. Shapes with no resolvable bounds (e.g.
 * placeholders that inherit position from the layout) are skipped.
 *
 * Useful for marquee-style region selection — e.g. "give me every
 * shape that lives in the right half of the slide" — without
 * iterating every shape and comparing bounds by hand.
 */
declare const findShapesInRect: (slide: SlideData, x: number, y: number, w: number, h: number) => ReadonlyArray<SlideShapeData>;
/**
 * Moves the shape so its center sits at the slide canvas center.
 * Reads the presentation's slide size, then sets the shape's
 * position to `(slideWidth/2 - shapeWidth/2, slideHeight/2 - shapeHeight/2)`.
 *
 * No-op when the shape has no bounds or the presentation has no
 * configured slide size.
 */
declare const centerShapeOnSlide: (shape: SlideShapeData) => void;
/**
 * `true` when two shapes' axis-aligned bounding boxes overlap.
 * Returns `false` when either shape has no bounds. Doesn't account
 * for rotation — uses the raw `<a:xfrm>` rectangle, not the
 * visual bounding box after rotation.
 *
 * Useful for collision detection in custom layout pipelines.
 */
declare const shapesOverlap: (a: SlideShapeData, b: SlideShapeData) => boolean;
/**
 * Sets both position and size in one call. Equivalent to calling
 * `setShapePosition` followed by `setShapeSize`, but commits the slide
 * just once.
 */
declare const setShapeBounds: (shape: SlideShapeData, bounds: ShapeBounds) => void;
//#endregion
//#region src/api/fn/shape-slide-read.d.ts
/**
 * Shapes on a slide, in document order with group children flattened.
 */
declare const getSlideShapes: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/**
 * Rebinds the slide to a different layout. The slide's own content
 * (shapes, text, geometry) is preserved verbatim; only the
 * `slideLayout` rel is updated so the reference desktop app re-renders with the new
 * layout's placeholder positions and theme.
 *
 * The new layout must already be a part of the package — pass one
 * returned by `getSlideLayouts(pres)` or `findSlideLayout(pres, name)`.
 */
declare const setSlideLayout: (slide: SlideData, layout: SlideLayoutData) => void;
/**
 * The slide layout this slide is bound to, or `null` if the slide has
 * no layout relationship.
 */
declare const getSlideLayout: (slide: SlideData) => SlideLayoutData | null;
/**
 * Returns the first placeholder shape with the given `type` (or `null`
 * if no match). Shapes whose `<p:ph>` omits an explicit type default to
 * `'body'` per ECMA-376 §19.7.10.
 */
declare const findSlidePlaceholder: (slide: SlideData, type: PlaceholderType) => SlideShapeData | null;
/**
 * Returns every placeholder shape on the slide whose text body is
 * empty. Useful for "spot the slots that still need filling" UIs
 * before a slide is published, and for validation hooks that warn
 * about empty slots.
 */
declare const findEmptyPlaceholders: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/**
 * Returns the union bounding box of a group of shapes, or `null`
 * when none of them have bounds. Useful for "select all and move
 * together" patterns where the caller needs a single rectangle
 * across the group.
 */
declare const getShapesBounds: (shapes: ReadonlyArray<SlideShapeData>) => ShapeBounds | null;
/**
 * Translates every shape in `shapes` by `(dxEmu, dyEmu)`. Useful
 * for "move this group of shapes 1cm right" patterns without
 * looping yourself. Shapes without bounds are skipped silently.
 */
declare const translateShapes: (shapes: ReadonlyArray<SlideShapeData>, dxEmu: number, dyEmu: number) => void;
/**
 * Returns every slide in the deck that has at least one empty
 * placeholder shape. Built on `findEmptyPlaceholders`. Useful for
 * "which slides still need editorial attention?" pre-publish
 * audits.
 */
declare const getSlidesWithEmptyPlaceholders: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns the first placeholder shape whose `<p:ph idx="...">`
 * matches `idx`, or `null` when none does. Real templates often
 * disambiguate same-type placeholders (e.g. two body slots) by
 * `idx`, so this is what you reach for when type-only lookup is
 * ambiguous.
 */
declare const findSlidePlaceholderByIdx: (slide: SlideData, idx: number) => SlideShapeData | null;
/**
 * Returns every placeholder shape with the given `type`. Useful for
 * "two-content" / "comparison" layouts where multiple body
 * placeholders share a type and the caller needs to fill them all.
 * Like `findSlidePlaceholder`, omitted `<p:ph type>` is treated as
 * `body` per ECMA-376 §19.7.10.
 */
declare const findSlidePlaceholders: (slide: SlideData, type: PlaceholderType) => ReadonlyArray<SlideShapeData>;
/**
 * First shape on the slide whose `cNvPr@name` matches `name`. Accepts
 * either a literal string (exact-equality) or a `RegExp` for pattern
 * matches — mirroring `findShapeByText`. Returns `null` when nothing
 * matches.
 */
declare const findShapeByName: (slide: SlideData, name: string | RegExp) => SlideShapeData | null;
/**
 * Returns the shape with the given OOXML internal id (`cNvPr@id`), or
 * `null` when no such shape exists. Shape ids are unique within a
 * slide; pair with `getShapeId` to round-trip references that arrive
 * from external XML (e.g. animations, hyperlinks).
 */
declare const findShapeById: (slide: SlideData, id: number) => SlideShapeData | null;
/**
 * Every shape on the slide whose `cNvPr@name` matches `name`. Accepts
 * either a literal string (exact-equality, case-sensitive) or a
 * `RegExp` for pattern matches — useful when template-cloned shapes
 * share a prefix (`'TextPlaceholder1'`, `'TextPlaceholder2'`, …).
 */
declare const findShapesByName: (slide: SlideData, name: string | RegExp) => ReadonlyArray<SlideShapeData>;
/**
 * First shape on the slide whose visible text matches `needle`
 * (substring or `RegExp`), or `null` when none does. Convenience
 * over `getSlideShapes(slide).find(...)` when the caller is hunting
 * for a label in a template ("find the box that says 'Q1'").
 */
declare const findShapeByText: (slide: SlideData, needle: string | RegExp) => SlideShapeData | null;
/**
 * Every shape on the slide whose visible text matches `needle`. Use
 * when more than one shape can share the same text (common with
 * cloned bullet templates) — multi-match variant of
 * `findShapeByText`.
 */
declare const findShapesByText: (slide: SlideData, needle: string | RegExp) => ReadonlyArray<SlideShapeData>;
/** Every shape on the slide of the given kind. */
declare const findShapesByKind: (slide: SlideData, kind: ShapeKind) => ReadonlyArray<SlideShapeData>;
/**
 * Every shape on the slide whose hyperlink target matches `url`
 * (substring or `RegExp`). Pairs the existing presentation-level
 * `findSlidesByHyperlink` for cases where the caller already has a
 * specific slide and wants the linking shapes inside it.
 */
declare const findShapesByHyperlink: (slide: SlideData, url: string | RegExp) => ReadonlyArray<SlideShapeData>;
/**
 * Every shape on the slide that carries any hyperlink (regardless of
 * target). Useful for "audit every clickable shape on this slide" —
 * counterpart to `findShapesByHyperlink(slide, url)`, which requires a
 * matching URL.
 */
declare const findShapesWithHyperlinks: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/**
 * Returns every shape on the slide whose `<a:prstGeom prst="…"/>`
 * token matches `preset`. Useful for diagram introspection: find all
 * `'leftArrow'` shapes for a workflow swap, replace every `'cloud'`
 * with `'rect'`, etc.
 *
 * Shapes without a preset (custGeom / pictures / charts / tables /
 * connectors / groups) are filtered out.
 */
declare const findShapesByPreset: (slide: SlideData, preset: string) => ReadonlyArray<SlideShapeData>;
/**
 * Returns the slide that owns `shape`. Useful when callers receive a
 * shape from an unfiltered walk (`getAllShapes`, `findShapeInPresentation`,
 * search results) and need to know which slide it's on.
 */
declare const getShapeSlide: (shape: SlideShapeData) => SlideData;
/**
 * Returns the shape's current XML element as a string. Diagnostic
 * sibling of `getSlideXmlString`; useful for snapshot tests, bug
 * reports, and before/after dumps during transformations.
 */
declare const getShapeXmlString: (shape: SlideShapeData) => string;
/**
 * Returns the 0-based document-order index of `shape` on its slide,
 * or `-1` when the shape is stale (e.g. after a `removeShape` that
 * rebuilt the slide's shape list).
 */
declare const getShapeIndex: (shape: SlideShapeData) => number;
/**
 * Walks every slide and returns the first shape whose name matches.
 * Accepts a literal string (exact equality) or a `RegExp` for
 * pattern matches — same shape as `findShapeByName` but deck-scoped.
 * Useful for "find the logo placeholder anywhere in the deck."
 */
declare const findShapeInPresentation: (pres: PresentationData, name: string | RegExp) => SlideShapeData | null;
/**
 * Replaces `{{key}}` tokens in every text-bearing shape on this slide.
 * Returns the number of substitutions performed.
 *
 * Tokens must fit within a single text run (see `replaceTokensInTree`
 * in `drawingml/`). Cross-run replacements aren't supported — use
 * `findSlidePlaceholder` + a setText path when the reference desktop app has
 * fragmented the run sequence.
 */
declare const replaceTokensInSlide: (slide: SlideData, tokens: Record<string, string>) => number;
//#endregion
//#region src/api/fn/shape-read-paint.d.ts
type ShapeFill = {
  readonly kind: 'solid';
  readonly color: string;
} | {
  readonly kind: 'gradient';
} | {
  readonly kind: 'pattern';
} | {
  readonly kind: 'image';
} | {
  readonly kind: 'background';
} | {
  readonly kind: 'none';
} | {
  readonly kind: 'inherit';
};
/**
 * Reads back the shape's stroke (`<a:ln>`). Returns:
 *
 *   - `{ kind: 'solid', color, widthEmu? }` for a solid-color outline.
 *   - `{ kind: 'gradient', widthEmu? }` for a gradient line; read its stops
 *     with `getShapeStrokeGradient`.
 *   - `{ kind: 'none' }` when an `<a:noFill>` sits inside `<a:ln>`.
 *   - `{ kind: 'inherit' }` when no `<a:ln>` is present.
 */
type ShapeStroke = {
  readonly kind: 'solid';
  readonly color: string;
  readonly widthEmu?: number;
} | {
  readonly kind: 'gradient';
  readonly widthEmu?: number;
} | {
  readonly kind: 'none';
} | {
  readonly kind: 'inherit';
};
/**
 * Convenience over `getShapeStroke(shape)`: returns the solid-
 * stroke color (`#RRGGBB` / `scheme:<token>`) or `null` when the
 * stroke is inherited / removed.
 */
declare const getShapeStrokeColor: (shape: SlideShapeData) => string | null;
/**
 * Convenience over `getShapeStroke(shape)`: returns the stroke
 * width in EMU when the stroke is solid and an explicit width is
 * set, or `null` otherwise.
 */
declare const getShapeStrokeWidth: (shape: SlideShapeData) => number | null;
/**
 * Returns the shape's stroke color resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the stroke isn't a solid color
 * (inherits / `noFill`) or when the color can't be resolved.
 *
 * Companion to `getShapeStrokeColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string — fine for round-tripping but
 * wrong for rendering, because the reference desktop app paints the *transformed*
 * color, not the base one.
 */
declare const getShapeStrokeColorResolved: (pres: PresentationData, shape: SlideShapeData) => string | null;
/**
 * Returns the opacity (`0`–`1`) of the shape's own solid outline, read from
 * the `<a:alpha>` / `<a:alphaMod>` / `<a:alphaOff>` children of its color
 * element, or `null` when the outline isn't a solid color or carries no
 * alpha transform (the reference desktop app draws it fully opaque). Companion to
 * `getShapeStrokeColorResolved`, which never carries the alpha channel.
 */
declare const getShapeStrokeOpacity: (shape: SlideShapeData, pres?: PresentationData) => number | null;
/**
 * Reads the stroke's line cap style — `'rnd'` (round), `'sq'` (square),
 * `'flat'`, or `null` when the attribute isn't set. Per ECMA-376
 * §20.1.2.3.10 (`ST_LineCap`).
 */
declare const getShapeStrokeCap: (shape: SlideShapeData, pres?: PresentationData) => "rnd" | "sq" | "flat" | null;
/**
 * Reads the stroke's line join style — `'round'` / `'bevel'` / `'miter'`,
 * or `null` when no explicit join element is present. Maps from the
 * three child-element variants `<a:round/>`, `<a:bevel/>`, `<a:miter/>`.
 */
declare const getShapeStrokeJoin: (shape: SlideShapeData, pres?: PresentationData) => "round" | "bevel" | "miter" | null;
/**
 * Reads the stroke's compound-line style (`<a:ln cmpd="…">`) — single,
 * double, triple, or thick/thin / thin/thick parallel lines. ECMA-376
 * §20.1.2.3.11 (`ST_CompoundLine`).
 */
declare const getShapeStrokeCompound: (shape: SlideShapeData, pres?: PresentationData) => "sng" | "dbl" | "thickThin" | "thinThick" | "tri" | null;
/**
 * Same as `getShapeStroke` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. First non-inherit
 * stroke layer wins.
 */
declare const getShapeStrokeEffective: (pres: PresentationData, shape: SlideShapeData) => ShapeStroke;
/**
 * Returns the full gradient (`stops`, direction, rotation) of a gradient line,
 * or `null` when the line is not one. Without `pres` only the shape's own
 * `<a:ln>` is read; with it, the line `getShapeStrokeEffective` resolves —
 * style matrix and placeholder cascade included — and every stop carries its
 * theme- and transform-resolved `resolvedColor`.
 */
declare const getShapeStrokeGradient: (shape: SlideShapeData, pres?: PresentationData) => ReadGradientFill | null;
declare const getShapeStroke: (shape: SlideShapeData) => ShapeStroke;
/**
 * Convenience over `getShapeFill(shape)`: returns the solid-fill
 * color string (`#RRGGBB` or `scheme:<token>`) when the shape has
 * one, or `null` otherwise. Use when the caller only cares about
 * the color and doesn't need to distinguish "inherit" / "no fill" /
 * "gradient" / "pattern" / "image" from each other.
 */
declare const getShapeFillColor: (shape: SlideShapeData) => string | null;
/**
 * Returns the shape's solid fill resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the fill isn't solid (gradient,
 * pattern, image, none, inherit) or when the color can't be resolved.
 *
 * Companion to `getShapeFillColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string. Renderers and exporters that
 * need the color the reference desktop app actually paints should call this.
 */
declare const getShapeFillColorResolved: (pres: PresentationData, shape: SlideShapeData) => string | null;
/**
 * Returns the opacity (`0`–`1`) of the shape's solid fill, read from
 * the `<a:alpha>` / `<a:alphaMod>` / `<a:alphaOff>` children of its color
 * element, or `null` when the fill isn't solid or carries no alpha
 * transform (the reference desktop app paints it fully opaque). Companion to
 * `getShapeFillColorResolved`, which never carries the alpha channel —
 * OOXML encodes color and alpha independently. Pass `pres` to resolve the
 * shape's theme fill reference when no direct fill is present.
 */
declare const getShapeFillOpacity: (shape: SlideShapeData, pres?: PresentationData) => number | null;
/**
 * Same as `getShapeFill` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. Returns the first
 * non-inherit fill found, or `{ kind: 'inherit' }` when neither layer
 * supplies one. Useful for renderers that want the actual fill the
 * placeholder will paint with.
 */
declare const getShapeFillEffective: (pres: PresentationData, shape: SlideShapeData) => ShapeFill;
declare const getShapeFill: (shape: SlideShapeData) => ShapeFill;
//#endregion
//#region src/api/fn/theme.d.ts
/**
 * The named color scheme on a presentation's theme. Each slot is a
 * `#RRGGBB` string — `sysClr` slots are flattened to their cached
 * `lastClr` value.
 */
interface PresentationTheme {
  readonly name: string;
  readonly dark1: string;
  readonly light1: string;
  readonly dark2: string;
  readonly light2: string;
  readonly accent1: string;
  readonly accent2: string;
  readonly accent3: string;
  readonly accent4: string;
  readonly accent5: string;
  readonly accent6: string;
  readonly hyperlink: string;
  readonly followedHyperlink: string;
}
/**
 * Returns the deck's color scheme as `#RRGGBB` strings, or `null` if
 * the package carries no theme. Each accent slot maps directly to the
 * `accent1`–`accent6` chart palette defaults.
 *
 * The deck's theme is the one its first slide master references. A
 * package usually holds more themes than that — notes and handout
 * masters carry their own, and Google Slides exports name the notes
 * theme `theme1.xml` ahead of the slide theme `theme2.xml` — so the
 * slide master relationship decides, not the part name. Multi-master
 * decks surface the first master's theme only; per-master theme lookup
 * will land if a concrete user need shows up.
 */
declare const getPresentationTheme: (pres: PresentationData) => PresentationTheme | null;
/**
 * The theme's font scheme, flattened to the typefaces the runs in a
 * deck inherit by default. `major*` is the heading font (slide
 * titles, chart titles); `minor*` is the body font.
 *
 * Each field carries the Latin / East-Asian / Complex-Script
 * typeface name as written on the theme. Empty values are
 * normalized to `null`.
 */
interface PresentationFonts {
  readonly majorLatin: string | null;
  readonly majorEastAsian: string | null;
  readonly majorComplexScript: string | null;
  readonly minorLatin: string | null;
  readonly minorEastAsian: string | null;
  readonly minorComplexScript: string | null;
}
/**
 * Returns the deck theme's font scheme, or `null` when the package
 * carries no theme. This is the presentation-level fallback; shape text
 * resolution uses `fontsFromThemeRoot` with the shape's owning master theme.
 */
declare const getPresentationFonts: (pres: PresentationData) => PresentationFonts | null;
/**
 * Overwrites the named slots of the deck theme's color
 * scheme (`<a:clrScheme>`), leaving every other slot untouched. Slot
 * values are `#RRGGBB` (or 3-digit `#RGB`) strings; every branded
 * `srgbClr`/`sysClr` slot is normalized to a plain `<a:srgbClr>` on
 * write, since a theme slot is never a scheme-color reference.
 *
 * As with `getPresentationTheme`, multi-master decks are branded via
 * their first slide master's theme only.
 *
 * Throws if the presentation has no theme part, or if a provided color
 * isn't a valid `#RRGGBB` string.
 */
declare const setPresentationTheme: (pres: PresentationData, theme: Partial<Omit<PresentationTheme, "name">> & {
  name?: string;
}) => void;
/**
 * Typeface overrides for {@link setPresentationFonts}. Unlike
 * `PresentationFonts` (the read side, where an empty typeface value
 * flattens to `null`), every field here is a plain string — the
 * underlying `<a:latin>` / `<a:ea>` / `<a:cs>` elements are mandatory on
 * a valid theme, so there's no "clear this typeface" operation.
 */
interface PresentationFontsInput {
  readonly majorLatin?: string;
  readonly majorEastAsian?: string;
  readonly majorComplexScript?: string;
  readonly minorLatin?: string;
  readonly minorEastAsian?: string;
  readonly minorComplexScript?: string;
}
/**
 * Overwrites the named typefaces of the deck theme's font scheme
 * (major = headings, minor = body), leaving unset fields untouched. As
 * with `setPresentationTheme`, only the first slide master's theme is
 * branded.
 *
 * Throws if the presentation has no theme part.
 */
declare const setPresentationFonts: (pres: PresentationData, fonts: PresentationFontsInput) => void;
//#endregion
//#region src/api/fn/shape-gradient-read.d.ts
/**
 * Returns the full gradient definition (`stops` + `angleDeg`) when the
 * shape's `<p:spPr>` carries an `<a:gradFill>`. Returns `null` for any
 * other fill kind, including `inherit` — the function does not walk the
 * layout / master cascade. Use `getShapeGradientFillEffective` for that.
 */
declare const getShapeGradientFill: (shape: SlideShapeData) => ReadGradientFill | null;
/**
 * Same as `getShapeGradientFill` but walks the layout → master
 * placeholder cascade when the shape carries no explicit fill choice.
 * Stop `resolvedColor` values include the theme, color map, and color transforms.
 * Returns the first gradient in the placeholder cascade, or `null` when an
 * explicit non-gradient fill masks the cascade or no layer defines one.
 *
 * Resolves literal gradients on the shape or its placeholder ancestors, and
 * the shape's own theme style-matrix fill reference. Style references on
 * placeholder ancestors are not resolved.
 */
declare const getShapeGradientFillEffective: (pres: PresentationData, shape: SlideShapeData) => ReadGradientFill | null;
//#endregion
//#region src/api/fn/shape-fill-stroke.d.ts
/**
 * Replace the preset or custom geometry of a shape, picture, or connector.
 * Picture geometry clips the image without modifying its bytes or crop.
 * Existing adjust guides are removed; use `setShapeAdjustValues` afterward
 * to customize the new preset. Position, size, fill, and effects are preserved.
 */
declare const setShapePreset: (shape: SlideShapeData, preset: PresetShape) => void;
/** Sets the shape's position in EMU. Companion to `setShapeSize`. */
declare const setShapePosition: (shape: SlideShapeData, x: Emu, y: Emu) => void;
/** Sets the shape's size in EMU. */
declare const setShapeSize: (shape: SlideShapeData, w: Emu, h: Emu) => void;
/**
 * Sets the shape's rotation in degrees (positive clockwise). Values are
 * normalized into `[0, 360)`; pass `0` to clear an existing rotation.
 */
declare const setShapeRotation: (shape: SlideShapeData, degrees: number) => void;
/**
 * Sets the shape's preset-geometry adjust values (`<a:prstGeom><a:avLst>`),
 * replacing any guides already authored. Keys are ECMA-376 guide names and
 * values are the raw guide numbers (rounded to integers). The companion
 * reader is {@link getShapeAdjustValues}.
 *
 * The common use is the corner radius of the `roundRect` preset, whose `adj`
 * guide runs `0..50000` (thousandths of a percent of the shorter side;
 * `16667` ≈ the reference desktop app's default, `0` = square corners, `50000` = fully
 * rounded). For example, `setShapeAdjustValues(shape, { adj: 5000 })` gives a
 * subtle 5% rounding.
 *
 * Throws when the shape has no preset geometry (`<a:prstGeom>`) — a custom or
 * inherited geometry has no adjust list to author.
 */
declare const setShapeAdjustValues: (shape: SlideShapeData, values: Readonly<Record<string, number>>) => void;
/** Sets the shape's flip flags. Properties default to current state when omitted. */
declare const setShapeFlip: (shape: SlideShapeData, options: {
  horizontal?: boolean;
  vertical?: boolean;
}) => void;
/**
 * Paints a shape with the slide background at its position, covering objects
 * behind it. This is different from a transparent fill. Only ordinary shapes
 * support the PresentationML `useBgFill` attribute; other kinds throw.
 */
declare const setShapeSlideBackgroundFill: (shape: SlideShapeData) => void;
/**
 * Sets a solid fill. A color string replaces the fill; an options object edits
 * only supplied properties and preserves existing opacity when changing color.
 * Opacity ranges from 0 (transparent) to 1 (opaque). Opacity-only edits require
 * an existing solid fill and preserve its theme reference and color transforms.
 */
declare const setShapeFill: (shape: SlideShapeData, color: Color | {
  color?: Color;
  opacity?: number;
}) => void;
/**
 * Sets a linear gradient fill on the shape. Stops must lie in `[0, 1]`;
 * `angleDeg` defaults to `90` (top → bottom).
 *
 * Example: red → blue top-to-bottom:
 *
 *   setShapeGradientFill(shape, {
 *     stops: [{ offset: 0, color: '#FF0000' }, { offset: 1, color: '#0000FF' }],
 *     angleDeg: 90,
 *   });
 */
declare const setShapeGradientFill: (shape: SlideShapeData, options: GradientFillOptions) => void;
/**
 * Sets a preset pattern fill on the shape (e.g. `pct50`, `dkUpDiag`).
 *
 * `foreground` is the pattern stroke color; `background` fills behind
 * the pattern. Both accept `#RRGGBB`, bare `RRGGBB`, or scheme tokens
 * (`accent1`, `bg1`, ...). Omitted settings preserve existing pattern XML,
 * including theme references and color transforms. A new pattern defaults to
 * `pct5`, foreground `accent1` and background `bg1`, as in the reference desktop app on Mac.
 */
declare const setShapePatternFill: (shape: SlideShapeData, options: Partial<PatternFillOptions>) => void;
/**
 * Reads back the pattern fill on a shape: returns the preset token
 * plus the foreground / background colors resolved against the theme.
 * With `preserveTheme`, untransformed scheme colors remain theme tokens.
 * Transformed colors still resolve to RGB to preserve their appearance.
 * Returns `null` when the shape has no `<a:pattFill>`.
 *
 * The preset is the literal `ST_PresetPatternVal` token from §20.1.10.49 —
 * e.g. `'pct50'`, `'dkUpDiag'`, `'cross'`, `'wave'`. Renderers can map it onto
 * an SVG `<pattern>` definition, and the result can be handed straight back to
 * `setShapePatternFill`. A missing or unrecognised `prst` reads as `'pct50'`,
 * which is what the reference desktop app paints when the attribute is absent.
 */
declare const getShapePatternFill: (pres: PresentationData, shape: SlideShapeData, options?: {
  readonly preserveTheme?: boolean;
}) => {
  preset: PatternPreset;
  foreground: string;
  background: string;
} | null;
/**
 * Sets a picture fill on the shape, embedding `bytes` as a new media
 * part and replacing any prior fill choice on the shape's `<p:spPr>`.
 *
 * The image stretches to fill the shape (`<a:stretch><a:fillRect/>`).
 * Format is detected from magic bytes; pass `options.format` to
 * override (useful for SVG or unusual extensions).
 *
 * Throws if the format can't be detected and isn't provided explicitly,
 * or if the shape kind doesn't carry a `<p:spPr>` (e.g. groups).
 */
declare const setShapeImageFill: (shape: SlideShapeData, bytes: Uint8Array, options?: {
  format?: ImageFormat;
}) => void;
/** Sets `<a:noFill>` on the shape, leaving it transparent. */
declare const setShapeNoFill: (shape: SlideShapeData) => void;
/**
 * Removes any fill choice from the shape; it then inherits its fill
 * from the layout / master placeholder it descends from.
 */
declare const clearShapeFill: (shape: SlideShapeData) => void;
/**
 * Updates outline color, width and/or opacity (0–1); omitted properties are
 * preserved. `fill: { kind: 'gradient', ... }` paints a gradient line instead
 * of `color` (the reference desktop app's Gradient line); it takes the same options as
 * `setShapeGradientFill` and is exclusive with `color` and `opacity`.
 */
declare const setShapeStroke: (shape: SlideShapeData, options: {
  color?: Color;
  widthEmu?: number;
  opacity?: number;
  fill?: LineFill;
}) => void;
/**
 * Sets the reference desktop app's Sketched style on the outline (`'curved'`, `'freehand'`
 * or `'scribble'`), or removes it with `null`. Stored as the 2021
 * `ask:lineSketchStyleProps` extension of `<a:ln>`, which older readers
 * ignore. A new sketch gets a seed derived from the shape id, so the same
 * deck always saves the same bytes.
 *
 * The reference desktop app writes the hand-drawn outline into the shape's geometry and keeps the
 * original beside the style; this library keeps the true geometry instead (and
 * restores it when editing a sketch the reference desktop app wrote), leaving the roughening to
 * the renderer.
 */
declare const setShapeStrokeSketch: (shape: SlideShapeData, sketch: LineSketch | null) => void;
/**
 * Reads the outline's sketched style, or `null` when it has none. Only the
 * shape's own `<a:ln>` is consulted.
 */
declare const getShapeStrokeSketch: (shape: SlideShapeData) => LineSketch | null;
/** Sets an explicit "no outline" on the shape. */
declare const setShapeNoStroke: (shape: SlideShapeData) => void;
/**
 * Reads back the shape's stroke dash style, or `null` if none. When `pres` is
 * supplied, a direct partial line is overlaid on the shape's `lnRef` style.
 */
declare const getShapeStrokeDash: (shape: SlideShapeData, pres?: PresentationData) => LineDash | null;
/**
 * Reads back the shape's arrowhead on one end of `<a:ln>`, or `null`
 * when no `<a:headEnd>` / `<a:tailEnd>` is present. When `pres` is supplied,
 * a direct partial line is overlaid on the shape's `lnRef` style.
 */
declare const getShapeStrokeArrow: (shape: SlideShapeData, end: "head" | "tail", pres?: PresentationData) => ArrowOptions | null;
/**
 * Sets the dash pattern for the shape's outline (`<a:prstDash>`). One
 * of ECMA-376's `ST_PresetLineDashVal` tokens:
 *
 *   `'solid'` | `'dot'` | `'dash'` | `'lgDash'` | `'dashDot'` |
 *   `'lgDashDot'` | `'lgDashDotDot'` | `'sysDash'` | `'sysDot'` |
 *   `'sysDashDot'` | `'sysDashDotDot'`
 *
 * Creates `<a:ln>` if absent. Pairs naturally with `setShapeStroke`:
 * users typically set a color + width first, then the dash.
 */
declare const setShapeStrokeDash: (shape: SlideShapeData, dash: LineDash) => void;
/**
 * Sets an arrowhead on one end of the shape's outline.
 *
 *   - `end: 'head'` writes `<a:headEnd>` (the start of the line).
 *   - `end: 'tail'` writes `<a:tailEnd>` (the end).
 *
 * Useful primarily on connector shapes added via `addSlideLine`.
 * `type: 'none'` clears the arrowhead.
 */
declare const setShapeStrokeArrow: (shape: SlideShapeData, end: "head" | "tail", options: ArrowOptions) => void;
/** Removes any outline override; the shape then inherits stroke from layout. */
declare const clearShapeStroke: (shape: SlideShapeData) => void;
/**
 * Sets the line-cap style on the shape's outline (`<a:ln cap="…">`).
 *
 *   - `'rnd'`  — rounded ends.
 *   - `'sq'`   — square ends extending past the endpoint.
 *   - `'flat'` — square ends flush at the endpoint (the OOXML default).
 *
 * Pass `null` to clear the attribute so the cap inherits the default.
 * Creates `<a:ln>` if absent.
 */
declare const setShapeStrokeCap: (shape: SlideShapeData, cap: "rnd" | "sq" | "flat" | null) => void;
/**
 * Sets the line-join style on the shape's outline. Picks one of the
 * three child-element variants of `<a:ln>`:
 *
 *   - `'round'` → `<a:round/>`
 *   - `'bevel'` → `<a:bevel/>`
 *   - `'miter'` → `<a:miter/>`
 *
 * Pass `null` to clear any prior join child so the shape inherits the
 * default. Creates `<a:ln>` if absent.
 */
declare const setShapeStrokeJoin: (shape: SlideShapeData, join: "round" | "bevel" | "miter" | null) => void;
/**
 * Sets the compound-line style on the shape's outline
 * (`<a:ln cmpd="…">`) — single, double, triple, or thick/thin variants.
 * ECMA-376 §20.1.10.31 `ST_CompoundLine`. Pass `null` to clear the
 * attribute. Creates `<a:ln>` if absent.
 */
declare const setShapeStrokeCompound: (shape: SlideShapeData, cmpd: "sng" | "dbl" | "thickThin" | "thinThick" | "tri" | null) => void;
//#endregion
//#region src/api/fn/shape-effects.d.ts
/** A shape's 3-D (`setShape3D`): the same DrawingML vocabulary as text 3-D. */
type Shape3D = Text3D;
/**
 * Read-back for `setShapeShadow` / `setShapeGlow`. Returns the kind
 * of effect currently on the shape's `<a:effectLst>`, or `null` when
 * none. Decodes the configured color + numeric parameters when
 * present.
 */
type ShapeEffect = {
  readonly kind: 'shadow';
  readonly color: string;
  readonly blurEmu: number;
  readonly offsetEmu: number;
  readonly angleDeg: number;
  readonly opacity?: number;
} | {
  readonly kind: 'glow';
  readonly color: string;
  readonly radiusEmu: number;
};
/**
 * Discriminated union covering every effect in
 * `CT_EffectStyleItem` (ECMA-376 §20.1.8.x) — outer shadow, inner
 * shadow, glow, reflection, soft-edge, blur. Returned in document
 * order so renderers can chain filters with the same composition
 * the reference desktop app applies.
 *
 * Lengths are EMU; angles are degrees clockwise from 3 o'clock;
 * opacity is a unit fraction (0..1) when the spec exposes one.
 */
type ShapeEffectAny = {
  readonly kind: 'outerShdw';
  readonly alignment?: ShadowOptions['alignment'];
  readonly rotateWithShape?: boolean;
  readonly color: string;
  readonly opacity?: number;
  readonly blurEmu: number;
  readonly distEmu: number;
  readonly angleDeg: number;
  readonly scaleX?: number;
  readonly scaleY?: number;
  readonly skewX?: number;
  readonly skewY?: number;
} | {
  readonly kind: 'innerShdw';
  readonly color: string;
  readonly opacity?: number;
  readonly blurEmu: number;
  readonly distEmu: number;
  readonly angleDeg: number;
} | {
  readonly kind: 'glow';
  readonly color: string;
  readonly opacity?: number;
  readonly radiusEmu: number;
} | {
  readonly kind: 'reflection';
  readonly opacity?: number;
  readonly startOpacity?: number;
  readonly startPosition?: number;
  readonly endPosition?: number;
  readonly fadeDirection?: number;
  readonly scaleX?: number;
  readonly alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
  readonly rotateWithShape?: boolean;
  readonly scaleY?: number;
  readonly skewX?: number;
  readonly skewY?: number;
  readonly blurEmu: number;
  readonly distEmu: number;
  readonly angleDeg: number;
} | {
  readonly kind: 'softEdge';
  readonly radiusEmu: number;
} | {
  readonly kind: 'blur';
  readonly radiusEmu: number;
};
declare const getShapeEffect: (shape: SlideShapeData) => ShapeEffect | null;
/**
 * Returns every effect attached to the shape's `<a:effectLst>` in
 * document order — outer shadow, inner shadow, glow, reflection,
 * soft edge, blur. Empty array when no effects apply.
 *
 * Companion to `getShapeEffect`, which is the v1 "first effect only"
 * helper. `getShapeEffects` is what renderers want because the reference desktop app
 * composes multiple effects in a single filter (shadow + glow, etc.).
 */
declare const getShapeEffects: (pres: PresentationData, shape: SlideShapeData) => readonly ShapeEffectAny[];
/**
 * Every shape on the slide whose `<a:effectLst>` carries an effect
 * of the given `kind` (`'outerShdw'`, `'innerShdw'`, `'glow'`,
 * `'reflection'`, `'softEdge'`, `'blur'`). Pure presence check — only
 * looks at the shape's own effect list, not the layout / master
 * cascade. Pair with `findShapesByEffect(pres, slide, 'softEdge')`
 * style call sites for visual-effect audits.
 */
declare const findShapesByEffect: (pres: PresentationData, slide: SlideData, kind: ShapeEffectAny["kind"]) => ReadonlyArray<SlideShapeData>;
/**
 * Same as `getShapeEffects` but walks the layout → master placeholder
 * cascade when the shape itself has no `<a:effectLst>`. Inherits
 * "all or nothing" — once any layer supplies an effect list, that
 * list is used; layers further down aren't merged in. This matches
 * the reference desktop app's behaviour (effect lists override rather than compose).
 */
declare const getShapeEffectsEffective: (pres: PresentationData, shape: SlideShapeData) => readonly ShapeEffectAny[];
/**
 * Sets an outer drop shadow on the shape. Defaults: black, 4pt blur,
 * 3pt offset, 45° (down-right). Pass `opacity` (0–1) to soften the
 * shadow. Replaces an existing drop shadow and leaves the shape's other
 * effects in place; `null` removes only the drop shadow and
 * `clearShapeEffects` removes every effect.
 */
declare const setShapeShadow: (shape: SlideShapeData, options?: ShadowOptions | null) => void;
/**
 * Sets an inner shadow (`<a:innerShdw>`) on the shape, replacing only a
 * prior inner shadow. Defaults: black, 4pt blur, 3pt offset, 45°. `null`
 * removes it.
 */
declare const setShapeInnerShadow: (shape: SlideShapeData, options: InnerShadowOptions | null) => void;
/**
 * Sets a glow around the shape. The radius is in EMU (default 5pt =
 * 63500). Replaces an existing glow and composes with a shadow, in the
 * order `CT_EffectList` states. `null` removes only the glow.
 */
declare const setShapeGlow: (shape: SlideShapeData, options: GlowOptions | null) => void;
/**
 * Sets a reflection (`<a:reflection>`) of the shape, replacing only a prior
 * reflection. A mirrored reflection below the shape is `scaleY: -1` with
 * `alignment: 'bl'`, fading from `startOpacity` to `opacity` by
 * `endPosition`. `null` removes it.
 */
declare const setShapeReflection: (shape: SlideShapeData, options: ReflectionOptions | null) => void;
/**
 * Softens the shape's edges (`<a:softEdge>`) by `radiusEmu`, replacing a
 * prior soft edge. `null` removes it.
 */
declare const setShapeSoftEdge: (shape: SlideShapeData, radiusEmu: number | null) => void;
/** Removes any effects (shadow / glow / future presets) from the shape. */
declare const clearShapeEffects: (shape: SlideShapeData) => void;
/**
 * Sets the shape's own 3-D — `<a:scene3d>` (camera, rotation, lighting) and
 * `<a:sp3d>` (bevels, depth, contour, material, distance from ground) in
 * `<p:spPr>`, what the reference desktop app's 3-D Format and 3-D Rotation write. Same
 * vocabulary as `setShapeText3D`, which puts it on the text body instead. A
 * field left out removes what it describes; settings this API does not model
 * (camera zoom, backdrop, ...) are kept while their element remains. `null`
 * removes both elements.
 */
declare const setShape3D: (shape: SlideShapeData, value: Shape3D | null) => void;
/**
 * Reads the 3-D on the shape's own `<p:spPr>` (see `setShape3D`), or `null`
 * when it has neither `<a:scene3d>` nor `<a:sp3d>`. A theme's effect style
 * is not consulted.
 */
declare const getShape3D: (shape: SlideShapeData) => ReadText3D | null;
//#endregion
//#region src/internal/drawingml/text-body-edit.d.ts
type TextCase = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle';
//#endregion
//#region src/api/fn/shape-text.d.ts
/**
 * Replaces the shape's visible text with `value`. Newlines start a new
 * paragraph by default; `newlines: 'break'` inserts paragraph-internal `<a:br>` instead.
 * By default, new paragraphs inherit the first existing run and
 * paragraph properties (font, color, size, alignment and bullets). The paragraph-end
 * format (`<a:endParaRPr>`) is not kept by default; author it with `setShapeParagraphs`.
 * Set `preserveFormatting` for incremental editing: unchanged prefix/suffix runs
 * and paragraphs retain their XML; inserted text inherits the insertion point's
 * format. Multiple disjoint changes should be applied separately to retain the
 * formatting between them. With `range`, `value` replaces exactly that UTF-16
 * selection and unaffected formatting is always preserved.
 * Pass `{ case: 'upper' | 'lower' | 'sentence' | 'title' | 'toggle' }` to apply
 * Change Case, as the reference desktop app does, to the whole text or to `options.range`.
 */
declare const setShapeText: (shape: SlideShapeData, value: string | {
  case: TextCase;
}, options?: {
  newlines?: "paragraph" | "break";
  bullets?: BulletStyle;
  preserveFormatting?: boolean;
  range?: {
    start: number;
    end: number;
  };
}) => void;
/**
 * Replaces the shape's text with a single field — text the reference desktop app fills in
 * when it opens the deck, rather than text the file states. `type` is an
 * ECMA-376 `ST_TextFieldType` token: `'slidenum'` for the slide's number,
 * `'datetime'` (and its `datetime1`…`datetime13` variants) for the current
 * date, `'footer'`, `'headerfooter'`, … Unrecognised tokens are written
 * through, since the list is open and renderers differ on what they honour.
 *
 * `options.text` is the cached value stored in `<a:t>`, which is what a reader
 * that does not evaluate fields shows — the reference desktop app overwrites it on open, so
 * it matters only for other consumers. The shape's first run's formatting is
 * carried onto the field, the way the reference desktop app keeps a placeholder's look when
 * it inserts one.
 *
 * This replaces the whole text body: a field placeholder holds the field and
 * nothing else, which is how the reference desktop app writes slide numbers and dates.
 */
declare const setShapeTextField: (shape: SlideShapeData, type: string, options?: {
  text?: string;
}) => void;
/**
 * Appends `value` to the shape's existing text on a new line. The
 * shape's existing run / paragraph formatting is preserved by
 * `setTextBody`; the new paragraph inherits the same template.
 *
 * Equivalent to `setShapeText(shape, getShapeText(shape) + '\n' + value)`,
 * minus the leading newline when there was no existing text.
 */
declare const appendShapeText: (shape: SlideShapeData, value: string) => void;
/**
 * Sets the vertical text anchor on the shape's text body
 * (`<a:bodyPr anchor="..."/>`). Choices map to ECMA-376 tokens:
 *
 *   - `'top'`    → `anchor="t"`
 *   - `'center'` → `anchor="ctr"`
 *   - `'bottom'` → `anchor="b"`
 *
 * The bodyPr is created if absent. Throws for non-text-bearing shape
 * kinds.
 */
type TextAnchor = 'top' | 'center' | 'bottom';
/**
 * Word wrap mode on a text body. `'square'` (the reference desktop app's default for
 * textboxes) wraps lines at the shape's width; `'none'` lets text
 * overflow horizontally.
 */
type TextWrap = 'none' | 'square';
/** Auto-fit mode on a text body. */
type TextAutoFit = 'none' | 'normal' | 'shape';
/**
 * Sets the text-body word-wrap mode.
 *
 *   - `'square'` writes `wrap="square"` — the reference desktop app's default for textboxes.
 *   - `'none'`   writes `wrap="none"`  — text can overflow horizontally.
 *
 * Throws for non-text-bearing shape kinds.
 */
declare const setShapeTextWrap: (shape: SlideShapeData, wrap: TextWrap) => void;
/** Reads back the bodyPr `wrap` attribute, or `null` when absent. */
declare const getShapeTextWrap: (shape: SlideShapeData) => TextWrap | null;
/**
 * Sets the text-body auto-fit mode:
 *
 *   - `'none'`   → `<a:noAutofit/>`
 *   - `'normal'` → `<a:normAutofit/>`   shrink text to fit the shape
 *   - `'shape'`  → `<a:spAutoFit/>`     grow the shape to fit text
 *
 * Replaces any prior auto-fit child on `<a:bodyPr>`. Throws for
 * non-text-bearing shape kinds.
 */
declare const setShapeTextAutoFit: (shape: SlideShapeData, mode: TextAutoFit) => void;
/**
 * Sets the 3-D on the shape's text body — `<a:scene3d>` and `<a:sp3d>` in
 * `<a:bodyPr>`, where the reference desktop app writes its text-art bevels. `value.scene`
 * writes the camera and light rig; `bevelTop`, `extrusionHeightEmu`,
 * `material` and `contourColor` write `<a:sp3d>` (replacing `<a:flatTx>`).
 * A field left out removes what it describes; settings this API does not
 * model (camera field of view, bottom bevel, extrusion color, ...) are kept
 * while their element remains. `null` removes both elements. Throws for
 * non-text-bearing shape kinds.
 */
declare const setShapeText3D: (shape: SlideShapeData, value: Text3D | null) => void;
/**
 * Sets the reference desktop app's Keep text flat (`<a:bodyPr><a:flatTx/>`): the text stays
 * flat, out of the shape's 3-D scene, when the shape is rotated in 3-D.
 * `<a:flatTx>` and the text body's own `<a:sp3d>` are one schema choice, so
 * keeping text flat removes a text bevel or extrusion set by
 * `setShapeText3D`. Throws for non-text-bearing shape kinds.
 */
declare const setShapeTextFlat: (shape: SlideShapeData, flat: boolean) => void;
/** Reads Keep text flat (see `setShapeTextFlat`) from the shape's own text body. */
declare const getShapeTextFlat: (shape: SlideShapeData) => boolean;
/**
 * Reads the 3-D on the shape's own text body (see `setShapeText3D`), or
 * `null` when its `<a:bodyPr>` has neither `<a:scene3d>` nor `<a:sp3d>`.
 * Inherited body properties are not consulted.
 */
declare const getShapeText3D: (shape: SlideShapeData) => ReadText3D | null;
/**
 * Reads back the bodyPr auto-fit child, or `null` when none is
 * present (the reference desktop app applies a layout-inherited default in that case).
 */
declare const getShapeTextAutoFit: (shape: SlideShapeData) => TextAutoFit | null;
/**
 * Reads the scale parameters the reference desktop app stores on `<a:normAutofit>`
 * once it has shrunk a text body to fit. Returns `null` if the body
 * doesn't carry `<a:normAutofit>` or the attributes are absent. Both
 * fields are unitless ratios in `[0, 1]`:
 *
 *   - `fontScale`     — multiply every run's font size by this. Default `1`.
 *   - `lnSpcReduction` — subtract from the line-height ratio. Default `0`.
 *
 * Companion to `getShapeTextAutoFit`. Renderers that want to match
 * the reference desktop app's actual on-screen text size apply these factors to the
 * authored font sizes; without them, every long title overflows.
 */
declare const getShapeTextAutoFitParams: (shape: SlideShapeData) => {
  fontScale: number;
  lnSpcReduction: number;
} | null;
/**
 * Reads back the vertical text anchor on the shape's `<a:bodyPr>`.
 * Maps the ECMA-376 tokens back to the public union:
 *
 *   `'t'` → `'top'`, `'ctr'` → `'center'`, `'b'` → `'bottom'`
 *
 * Returns `null` when the bodyPr is absent or has no anchor attribute.
 */
declare const getShapeTextAnchor: (shape: SlideShapeData) => TextAnchor | null;
/**
 * Reads back the internal margins of the shape's text frame. Sides
 * that are absent in the XML default to `null` (the reference desktop app applies
 * its built-in default for the missing side).
 */
/**
 * Reads the multi-column layout on a text body — `<a:bodyPr
 * numCol="N" spcCol="EMU"/>`. Returns `null` when columns aren't
 * configured (the default single column). `gapEmu` is the
 * inter-column gap in EMU; omitted when `<a:bodyPr>` has no
 * `spcCol` attribute.
 */
declare const getShapeTextColumns: (shape: SlideShapeData) => {
  count: number;
  gapEmu?: number;
} | null;
/**
 * Sets the multi-column layout on the shape's text body — writes
 * `<a:bodyPr numCol="N" [spcCol="EMU"]/>`. Pass `null` to clear both
 * attributes so the text body inherits its column settings (one column
 * when no layout/master supplies them). `count` must be in `1..16`;
 * use `1` to override inherited columns explicitly. `gapEmu`, when omitted,
 * removes any prior `spcCol`. Throws for non-text-bearing shape kinds.
 */
declare const setShapeTextColumns: (shape: SlideShapeData, columns: {
  count: number;
  gapEmu?: number;
} | null) => void;
/**
 * Reads the shape's text-body rotation from `<a:bodyPr rot="N"/>`.
 * `rot` is stored in 60000ths of a degree (OOXML angle units); the
 * returned value is in degrees. Positive values rotate clockwise per
 * the reference desktop app's convention. Returns `null` when the attribute is
 * absent.
 *
 * Distinct from the shape's own `<p:xfrm rot=…>` (the geometry
 * rotation surfaced via the shape's `rotation`); `bodyPr rot` rotates
 * the text body inside the shape without rotating the shape itself.
 */
declare const getShapeTextBodyRotationDeg: (shape: SlideShapeData) => number | null;
/**
 * Sets the shape's text-body rotation (`<a:bodyPr rot="N"/>`), measured
 * in degrees. Positive rotates clockwise per the reference desktop app's convention.
 * Passing `null` clears the attribute so the shape inherits the default
 * (`0`). Throws for non-text-bearing shape kinds.
 *
 * Companion to `setShapeRotation`, which rotates the *whole* shape
 * via `<p:xfrm rot>`. `bodyPr rot` rotates only the text inside.
 */
declare const setShapeTextBodyRotationDeg: (shape: SlideShapeData, rotationDeg: number | null) => void;
/**
 * Reads the shape's text-direction token from `<a:bodyPr vert="…"/>`.
 * Per ECMA-376 §17.18.93 `ST_TextVerticalType`:
 *
 *   - `horz` — default left-to-right, top-to-bottom (returns `null`).
 *   - `vert` — 90° rotation, lines run top-to-bottom, columns right-to-left.
 *   - `vert270` — 270° rotation, lines bottom-to-top, columns left-to-right.
 *   - `wordArtVert` — characters not rotated, stacked vertically.
 *   - `eaVert` — East-Asian vertical: characters upright, columns right-to-left.
 *   - `mongolianVert` — Mongolian: rotated 90°, columns left-to-right.
 *   - `wordArtVertRtl` — upright characters stacked vertically, columns right-to-left.
 *
 * Returns `null` when the attribute is absent or set to the default
 * `horz`.
 */
declare const getShapeTextDirection: (shape: SlideShapeData) => "vert" | "vert270" | "wordArtVert" | "eaVert" | "mongolianVert" | "wordArtVertRtl" | null;
/**
 * Sets the shape's text-direction via `<a:bodyPr vert="…"/>`. See
 * `getShapeTextDirection` for the meaning of each value. Passing `null`
 * clears the attribute, restoring layout/master inheritance. `'horz'` writes
 * an explicit horizontal override. Throws for non-text-bearing shape kinds.
 */
declare const setShapeTextDirection: (shape: SlideShapeData, direction: "horz" | "vert" | "vert270" | "wordArtVert" | "eaVert" | "mongolianVert" | "wordArtVertRtl" | null) => void;
declare const getShapeTextMargins: (shape: SlideShapeData) => {
  readonly left: number | null;
  readonly top: number | null;
  readonly right: number | null;
  readonly bottom: number | null;
} | null;
/**
 * Resolves the effective `<a:bodyPr>` properties — anchor, wrap, vertical
 * direction, inset margins, columns and autofit — by walking the layout / master cascade
 * the same way `getShapeRunFormatEffective` walks rPr. Returns the
 * innermost value that the cascade supplies, or `null` for properties
 * neither the shape nor any inherited placeholder authors.
 *
 * Companion to `getShapeTextAnchor` / `getShapeTextWrap` /
 * `getShapeTextDirection` / `getShapeTextMargins`, which only report the
 * literal value on the shape itself.
 */
declare const getShapeBodyPrEffective: (pres: PresentationData, shape: SlideShapeData) => {
  anchor: TextAnchor | null;
  anchorCentered: boolean | null;
  autoFit: TextAutoFit | null;
  autoFitParams: ReturnType<typeof getShapeTextAutoFitParams>;
  columns: ReturnType<typeof getShapeTextColumns>;
  wrap: TextWrap | null;
  vert: ReturnType<typeof getShapeTextDirection>;
  margins: {
    left: number | null;
    top: number | null;
    right: number | null;
    bottom: number | null;
  };
};
/**
 * Sets vertical anchoring. `centered` centers the text block without changing
 * paragraph alignment (the reference desktop app's Top / Middle / Bottom Centered options).
 * Omit it to preserve centering; pass null to restore inherited centering.
 */
declare const setShapeTextAnchor: (shape: SlideShapeData, anchor: TextAnchor, options?: {
  centered?: boolean | null;
}) => void;
/**
 * Sets the internal margins of the shape's text frame in EMU. Each
 * side is independent; omitted sides keep their current value (or the
 * layout-inherited default when the attribute is absent).
 *
 * The reference desktop app's defaults for a textbox: left/right 91440 (0.1in),
 * top/bottom 45720 (0.05in).
 *
 *   setShapeTextMargins(shape, { left: 0, right: 0 }); // flush-left text
 */
declare const setShapeTextMargins: (shape: SlideShapeData, margins: {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}) => void;
/** Sets the bullet style on every paragraph in the shape's text body. */
declare const setShapeBulletStyle: (shape: SlideShapeData, style: BulletStyle) => void;
/** Sets the horizontal alignment of every paragraph in the shape's text. */
declare const setShapeAlignment: (shape: SlideShapeData, align: ParagraphAlignment) => void;
/**
 * Applies `format` to every run in the shape's text. Run-property
 * attributes not addressed by `format` are preserved, so partial
 * updates compose. Optional `range` selects UTF-16 offsets in `getShapeText`,
 * with an exclusive end. Paragraph separators and line breaks each count as one
 * character. Empty ranges do nothing; invalid or split-surrogate boundaries throw.
 * Partially selected fields become literal runs; fully selected fields stay fields.
 * Pass a zero-based `paragraphEnd` instead of `range` to update a paragraph's end mark
 * without rebuilding its runs or paragraph properties.
 * With `reset`, clears direct visual run properties before applying `format`,
 * restoring inherited fonts and appearance while keeping links and language.
 * Without either target, also clears text-body and paragraph run-format defaults.
 * Blank autoshapes receive a text body and paragraph-end formatting for future input.
 */
declare const setShapeTextFormat: (shape: SlideShapeData, format: TextFormat, options?: {
  range?: {
    start: number;
    end: number;
  };
  reset?: boolean;
  paragraphEnd?: never;
} | {
  paragraphEnd: number;
  reset?: boolean;
  range?: never;
}) => void;
/**
 * Replaces the shape's text with explicitly structured paragraphs, each
 * carrying its own runs and per-run formats. `setShapeText` yields one run
 * per paragraph; use this when a paragraph mixes formats (a bold lead-in
 * followed by plain text, for example). Read back with
 * `getShapeParagraphElements`.
 * Pass `{ source, range? }` to copy existing paragraph XML, including fields,
 * run formatting, bullets and remapped relationships. The optional range uses
 * UTF-16 offsets with an exclusive end; omitted range copies all source text.
 * Pass an array of targets and `{ source, ranges }` to distribute ranges in one
 * batch. All ranges are read before any target changes, including when the source
 * is a target. Each target must have a corresponding range.
 * Pass `{ sources }` to concatenate complete paragraph sequences in order,
 * preserving fields, paragraph properties and relationships across slides.
 * A single target with `{ source, ranges }` concatenates those ranges in order.
 * Target body properties and list styles remain unchanged.
 */
declare function setShapeParagraphs(shape: SlideShapeData, paragraphs: ReadonlyArray<ParagraphSpec> | {
  source: SlideShapeData;
  range?: {
    start: number;
    end: number;
  };
} | {
  sources: ReadonlyArray<SlideShapeData>;
} | {
  source: SlideShapeData;
  ranges: ReadonlyArray<{
    start: number;
    end: number;
  }>;
}): void;
declare function setShapeParagraphs(shapes: ReadonlyArray<SlideShapeData>, paragraphs: {
  source: SlideShapeData;
  ranges: ReadonlyArray<{
    start: number;
    end: number;
  }>;
}): void;
//#endregion
//#region src/api/fn/shape-color.d.ts
/**
 * Resolves a DrawingML color element (`<a:srgbClr>` / `<a:schemeClr>` /
 * `<a:sysClr>` / `<a:prstClr>`) with all its `<a:lumMod>` / `<a:tint>` /
 * `<a:shade>` / `<a:satMod>` etc. transform children applied. Returns
 * `null` when the color is a scheme token and no theme is supplied to
 * resolve it.
 *
 * Exposed because both run-format and fill-format code paths need to
 * apply the same transform pipeline; keeping a single implementation
 * means future spec-coverage additions only have to land in one place.
 */
declare const resolveDrawingColor: (colorEl: XmlElement, theme: PresentationTheme | null, clrMap?: Readonly<Record<string, string>> | null) => string | null;
/**
 * Resolves the opacity a DrawingML color element carries through its
 * `<a:alpha>` / `<a:alphaMod>` / `<a:alphaOff>` children (ECMA-376
 * §20.1.2.3.1–3) to a `0`–`1` fraction, applied in document order from a
 * fully opaque base. Returns `null` when the element has no alpha
 * transform — the reference desktop app paints that opaque, but callers can still tell
 * "unspecified" apart from an explicit `1`.
 *
 * Kept separate from `resolveDrawingColor` because alpha never changes the
 * `#RRGGBB`; OOXML encodes the two independently and renderers emit the
 * opacity next to the color (`fill-opacity` / `stroke-opacity`).
 */
declare const resolveDrawingColorOpacity: (colorEl: XmlElement) => number | null;
/**
 * Reads back the format of a single run. Returns `null` when the run
 * has no `<a:rPr>` (it inherits its format from the paragraph /
 * layout / master). Boolean attributes that are explicitly `"0"`
 * decode to `false`.
 *
 * Use `getShapeRunFormatEffective` if you want the resolved format
 * after walking the placeholder / lstStyle / master inheritance chain.
 */
declare const getShapeRunFormat: (shape: SlideShapeData, paragraphIndex: number, runIndex: number) => ReadTextFormat | null;
//#endregion
//#region src/api/fn/shape-paragraph.d.ts
/**
 * Resolves a run's effective character properties by walking the
 * ECMA-376 §21.1.2.4.7 inheritance chain — run rPr →
 * pPr defRPr → text-body lstStyle → shape Quick Style → layout placeholder
 * lstStyle → master placeholder lstStyle + master txStyles → owning theme
 * fontScheme.
 *
 * Each property (font, size, color, bold, italic, underline) is
 * resolved independently: the innermost layer that supplies a value
 * wins for that one property.
 *
 * Returns a non-null `TextFormat`; fields the cascade couldn't
 * resolve are simply absent (the renderer falls back to placeholder
 * defaults).
 *
 * Use `getShapeRunFormat` if you only want the literal `<a:rPr>` on
 * the run without inheritance. Numeric indices count only regular runs;
 * `{ fieldIndex }` and `{ breakIndex }` count fields and line breaks separately;
 * null resolves the paragraph end mark.
 * `inheritanceSource` supplies the original
 * placeholder and slide context when reading a detached editing preview;
 * paragraph and run properties are still read from `shape`.
 */
declare const getShapeRunFormatEffective: (pres: PresentationData, shape: SlideShapeData, paragraphIndex: number, runIndex: number | null | {
  readonly fieldIndex: number;
} | {
  readonly breakIndex: number;
}, options?: {
  inheritanceSource?: SlideShapeData;
}) => ReadTextFormat;
/** A custom paragraph tab stop. */
interface ParagraphTabStop {
  /** Position relative to the left margin, in EMU. */
  positionEmu: number;
  alignment: 'left' | 'center' | 'right' | 'decimal';
}
/** Effective bullet detail returned by `getParagraphPropertiesEffective`. */
interface ParagraphBulletDetail {
  color: string | null;
  colorFollowText: boolean;
  sizePct: number | null;
  sizePts: number | null;
  sizeFollowText: boolean;
  font: string | null;
  fontFollowText: boolean;
}
/** Effective paragraph properties returned by `getParagraphPropertiesEffective`. */
interface ParagraphProperties {
  /** Custom tab stops; an empty list explicitly clears inherited stops. */
  tabStops?: ParagraphTabStop[];
  /** Distance between automatic tab stops, in EMU. */
  defaultTabSizeEmu?: number;
  /** Horizontal alignment per `ParagraphAlignment`. */
  align: ParagraphAlignment | null;
  /** East Asian line-breaking rules; absent when not authored or inherited. */
  asianLineBreak?: boolean;
  /** Allow a Latin word to break across lines. */
  latinLineBreak?: boolean;
  /** Allow punctuation to extend beyond the text margin. */
  hangingPunctuation?: boolean;
  /** Vertical alignment of differently sized characters within a line. */
  fontAlignment?: 'auto' | 'top' | 'center' | 'baseline' | 'bottom';
  /** Outline level (0..8). 0 = top-level paragraph. */
  level: number;
  /** Left indent in EMU. */
  marL: number | null;
  /** Right indent in EMU. */
  marR: number | null;
  /** First-line indent in EMU; negative for hanging indents. */
  indent: number | null;
  /** Line spacing — either a percent multiplier or a fixed point value. */
  lineSpacing: {
    readonly kind: 'pct';
    readonly value: number;
  } | {
    readonly kind: 'pts';
    readonly value: number;
  } | null;
  /** Space before the paragraph in points. */
  spcBefPts: number | null;
  /** Space after the paragraph in points. */
  spcAftPts: number | null;
  /** Right-to-left paragraph (`<a:pPr rtl="1"/>`). */
  rtl: boolean | null;
  /**
   * Bullet resolved through the cascade (`<a:buChar>` / `<a:buAutoNum>` /
   * `<a:buNone>`). Master `bodyStyle` levels author the reference desktop app's default
   * "•", so body placeholders get their bullet from here even when no
   * slide-level `<a:pPr>` carries one.
   */
  bullet: BulletStyle | null;
  /** Bullet color, size, and font after paragraph/layout/master inheritance. */
  bulletDetail?: ParagraphBulletDetail;
}
/**
 * Resolves a paragraph's effective properties by walking the same
 * inheritance chain `getShapeRunFormatEffective` uses, but for the
 * paragraph-level surface:
 *
 *   - alignment, indent (left / right / first-line), line spacing,
 *     paragraph spacing (before / after), rtl.
 *
 * Each property is resolved independently; the innermost layer that
 * sets it wins. Fields the cascade can't resolve come through as `null`
 * so renderers know to fall back to their own defaults.
 *
 * Companion to `getParagraphAlignment` / `getParagraphLineSpacing` /
 * `getParagraphIndent` / `getParagraphSpacing`, which only surface the
 * literal `<a:pPr>` and skip the layout / master cascade.
 * `inheritanceSource` preserves the original placeholder and slide context
 * when resolving a detached shape preview. Table cells use their own text body.
 */
declare const getParagraphPropertiesEffective: (pres: PresentationData, shape: SlideShapeData | TableCellData, paragraphIndex: number, options?: {
  inheritanceSource?: SlideShapeData;
}) => ParagraphProperties;
/**
 * Applies `format` to a single run. Run-property attributes not
 * addressed by `format` are preserved — partial updates compose.
 *
 * Example: bold the second word of the first paragraph:
 *
 *   setShapeRunFormat(shape, 0, 1, { bold: true, color: '#FF0000' });
 */
declare const setShapeRunFormat: (shape: SlideShapeData, paragraphIndex: number, runIndex: number, format: TextFormat) => void;
/**
 * Reads the external URL the first run in the shape's text-body links
 * to (set via `setShapeHyperlink`). Returns `null` when no run carries
 * an `<a:hlinkClick r:id=…/>` or the rId resolves to a non-hyperlink
 * target.
 */
declare const getShapeHyperlink: (shape: SlideShapeData) => string | null;
/**
 * Sets an external hyperlink on every run in the shape's text. Allocates
 * (or reuses) a `hyperlink` relationship on the slide's `.rels`. Pass
 * `null` to clear. Optional `range` limits the change to UTF-16 offsets in
 * getShapeText, preserving text and formatting outside the selection.
 */
declare const setShapeHyperlink: (shape: SlideShapeData, url: string | null, tooltip?: string, options?: {
  range?: {
    start: number;
    end: number;
  };
}) => void;
//#endregion
//#region src/api/fn/comments.d.ts
/**
 * Every author known to the package's `commentAuthors.xml`.
 * Returns an empty array when no author list exists.
 */
declare const getCommentAuthors: (pres: PresentationData) => ReadonlyArray<CommentAuthor>;
/**
 * Returns every slide that has at least one comment by the given
 * author name. Accepts a literal string (exact-equality) or a
 * `RegExp` for pattern matches. Sibling of `findCommentsByAuthor`
 * (which returns the comments themselves).
 */
declare const findSlidesWithCommentsByAuthor: (pres: PresentationData, authorName: string | RegExp) => ReadonlyArray<SlideData>;
/**
 * Returns every distinct author who has at least one comment
 * anywhere in the deck. Deduplicates by author id; preserves
 * first-seen order. Differs from `getCommentAuthors(pres)`, which
 * surfaces every author registered in `commentAuthors.xml` even
 * when no comments reference them.
 */
declare const getPresentationCommenters: (pres: PresentationData) => ReadonlyArray<CommentAuthor>;
/**
 * Histogram of comment counts by author display name across the whole
 * deck. Useful for "who reviewed this deck the most?" audits.
 * Authors with the same display name (a real-world case for shared
 * mailbox identities) get merged into the same bucket; pair with
 * `getPresentationCommenters` when you need to keep authors with
 * identical names separate by `id`.
 */
declare const getPresentationCommentCountsByAuthor: (pres: PresentationData) => Readonly<Record<string, number>>;
/**
 * Dense histogram of comment counts by 0-based slide index. Every
 * slide in the deck appears as an element (count `0` when the slide
 * has no comments), so the array shape is dense — handy for charting
 * comment density per slide without re-indexing.
 */
declare const getPresentationCommentCountsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * Looks up a `CommentAuthor` from `commentAuthors.xml` by display
 * name. Accepts a literal string (exact equality) or a `RegExp` for
 * pattern matches. Returns `null` when no author matches. Sibling of
 * `findCommentsByAuthor` — the latter returns the matching comments;
 * this returns the author handle for downstream metadata reads (id,
 * initials, color).
 */
declare const findCommentAuthorByName: (pres: PresentationData, authorName: string | RegExp) => CommentAuthor | null;
/**
 * Returns every comment whose author name matches `authorName`
 * across every slide in the deck. Accepts a literal string (exact-
 * equality) or a `RegExp` for pattern matches — useful for reviewer-
 * specific filters ("show me all of Alice's notes" / `/^review-bot/`).
 */
declare const findCommentsByAuthor: (pres: PresentationData, authorName: string | RegExp) => ReadonlyArray<SlideCommentData>;
/**
 * Returns every comment whose text matches `needle` (substring or
 * `RegExp`) across the whole deck. Sibling of `findCommentsByAuthor`
 * — useful for "find every comment that mentions X" reviewer flows.
 */
declare const findCommentsByText: (pres: PresentationData, needle: string | RegExp) => ReadonlyArray<SlideCommentData>;
/**
 * Returns every comment whose `@dt` timestamp is **strictly after**
 * `since` (an ISO-8601 string or `Date`). Comments missing a `dt`
 * are skipped. Sibling of `findCommentsByText` / `findCommentsByAuthor`
 * — useful for "what's new since my last review?" surfaces.
 */
declare const findCommentsAfter: (pres: PresentationData, since: string | Date) => ReadonlyArray<SlideCommentData>;
/**
 * Returns every comment whose `@dt` timestamp is **strictly before**
 * `until` (an ISO-8601 string or `Date`). Comments missing a `dt`
 * are skipped. Sibling of `findCommentsAfter`.
 */
declare const findCommentsBefore: (pres: PresentationData, until: string | Date) => ReadonlyArray<SlideCommentData>;
/**
 * Returns every comment carrying a parseable `@dt`, sorted oldest
 * to newest. Comments without a date are omitted. Use `.at(0)` for
 * the oldest and `.at(-1)` for the newest.
 */
declare const getCommentsSortedByDate: (pres: PresentationData) => ReadonlyArray<SlideCommentData>;
/**
 * Returns the distinct authors who commented on this slide, in
 * first-seen order. Dedupes by author id. Sibling of
 * `getPresentationCommenters` for a slide-scoped reviewer roster.
 */
declare const getSlideCommentAuthors: (slide: SlideData) => ReadonlyArray<CommentAuthor>;
/**
 * Returns every comment attached to the slide, with the author already
 * resolved. Both formats are read: the ECMA-376 `<p:cm>` list first, then
 * the modern threads. The list is read-only — use `addSlideComment` /
 * `removeSlideComment` to mutate.
 */
declare const getSlideComments: (slide: SlideData) => ReadonlyArray<SlideCommentData>;
/** Options shared by both comment formats. */
interface AddCommentOptions {
  author: {
    name: string;
    initials?: string;
  };
  text: string;
  position?: CommentPosition | null;
  date?: Date;
  /** Existing comment on this slide to reply to. */
  replyTo?: SlideCommentData;
}
/**
 * Adds a comment to the slide. Returns the new comment handle.
 *
 * Which format it is written in follows the file rather than the caller: a
 * reply joins its own thread, and a new thread goes where the slide — or,
 * failing that, the deck — already keeps its comments. Only a deck with no
 * comments at all starts an ECMA-376 `<p:cm>` list, which everything can
 * read.
 *
 * Author handling: if an author with the given `name`+`initials` already
 * exists in the relevant author list, the existing record is reused (and, in
 * the legacy list, its `lastIdx` is bumped). Otherwise a new author is
 * allocated. `initials` defaults to the first character of `name`.
 *
 * `position` is the comment pin, in EMU. The legacy schema makes `<p:pos>`
 * mandatory, so a legacy comment added without one is pinned to the slide's
 * origin rather than written without the element; the modern schema makes it
 * optional, so a modern comment added without one has no pin.
 *
 * `date` defaults to the current time.
 */
declare const addSlideComment: (slide: SlideData, opts: AddCommentOptions) => SlideCommentData;
/** Updates comment text while retaining author, date, position, and extension XML. */
declare const setCommentText: (comment: SlideCommentData, text: string) => void;
/**
 * Removes the comment and all its replies from its slide's comments part. If the comment
 * was the last one on the slide, the comments part and the
 * slide → comments rel are also removed. The author entry in
 * `commentAuthors.xml` is left intact (an author may have comments on
 * other slides).
 */
declare const removeSlideComment: (comment: SlideCommentData) => void;
/**
 * Strips every comment from every slide in the deck. Returns the
 * number of comments removed. Built on `writeCommentsForSlide`,
 * so each slide's legacy comment part is rewritten with an empty
 * list. The `commentAuthors.xml` registry is left intact for any
 * caller that still needs author identity.
 *
 * Useful as a sanitizer before sharing a draft externally — pairs
 * with `clearAllSlideNotes` for a "remove reviewer chatter" pass.
 */
declare const clearAllSlideComments: (pres: PresentationData) => number;
/**
 * Slide-scoped sibling of `clearAllSlideComments`. Removes every
 * comment on the given slide and returns the number removed.
 */
declare const clearSlideComments: (slide: SlideData) => number;
declare const getCommentAuthor: (comment: SlideCommentData) => CommentAuthor;
declare const getCommentText: (comment: SlideCommentData) => string;
declare const getCommentDate: (comment: SlideCommentData) => string | null;
declare const getCommentPosition: (comment: SlideCommentData) => CommentPosition | null;
/**
 * Returns the slide that owns this comment. Counterpart to
 * `getShapeSlide(shape)` — handy when filtering across the deck
 * with `findCommentsByAuthor` and the caller needs to know which
 * slide each hit came from.
 */
declare const getCommentSlide: (comment: SlideCommentData) => SlideData;
/** Parent handle from the same read, or null for roots and missing parents. */
declare const getCommentParent: (comment: SlideCommentData) => SlideCommentData | null;
/**
 * Which of the two comment formats this handle came from. `'legacy'` is the
 * ECMA-376 `<p:cm>` list every reader understands; `'modern'` is the
 * `<p188:cm>` thread the reference desktop app writes from 2021 on.
 */
declare const getCommentFormat: (comment: SlideCommentData) => "legacy" | "modern";
/**
 * Whether the thread has been resolved. `null` where the question does not
 * apply or cannot be answered: a legacy comment, which has nowhere to record
 * it, or a modern one whose file states a status this schema version does not
 * define.
 */
declare const getCommentStatus: (comment: SlideCommentData) => CommentStatus | null;
/**
 * Resolves a thread, reopens it (`'active'`), or closes it. Replies carry a
 * status of their own, so this works on any modern comment.
 *
 * Throws for a legacy comment: the ECMA-376 schema has no place to put it,
 * and quietly doing nothing would leave the caller believing a thread had
 * been resolved.
 */
declare const setCommentStatus: (comment: SlideCommentData, status: CommentStatus) => void;
//#endregion
//#region src/api/fn/shape-image.d.ts
/**
 * How an image fills its target box:
 *
 *   - `'fill'` — stretch to the exact `w × h`, ignoring aspect ratio
 *     (the historical behavior, and the default for back-compat).
 *   - `'contain'` — scale to fit inside the box preserving aspect ratio,
 *     then center the result. The box's empty margins stay transparent.
 */
type ImageFit = 'fill' | 'contain';
/**
 * Replaces only this picture's media with `bytes`, detaching shared media
 * or relationships when necessary. Unshared same-format media is updated
 * in place. The geometry — crop, transform — is preserved.
 *
 * Pass `options.fit: 'contain'` to re-fit the picture's extent to the
 * replacement image's aspect ratio, inscribed and centered in the shape's
 * current `(off, ext)` box. The default `'fill'` leaves the extent as-is
 * (the historical behavior). When the new image's natural size can't be
 * measured (non-PNG/JPEG), `'contain'` is a no-op rather than an error.
 */
declare const setShapeImage: (shape: SlideShapeData, bytes: Uint8Array, options?: {
  format?: ImageFormat;
  fit?: ImageFit;
}) => void;
//#endregion
//#region src/api/fn/shape-click-action.d.ts
/** What clicking the shape should do. */
type ShapeClickAction = {
  readonly kind: 'url';
  readonly url: string;
} | {
  readonly kind: 'slide';
  readonly slide: SlideData;
} | {
  readonly kind: 'nextSlide';
} | {
  readonly kind: 'prevSlide';
} | {
  readonly kind: 'firstSlide';
} | {
  readonly kind: 'lastSlide';
} | {
  readonly kind: 'lastSlideViewed';
} | {
  readonly kind: 'endShow';
} | {
  readonly kind: 'customShow';
  readonly id: number;
  readonly returnToShow: boolean;
};
/**
 * Reads the click action attached to the shape's cNvPr, or `null` if
 * none. Mirrors `setShapeClickAction`:
 *
 *   - `{ kind: 'url', url }`     — `hyperlink` rel + targetMode=External
 *   - `{ kind: 'slide', slide }` — `slide` rel + `ppaction://hlinksldjump`
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *     — preset show-navigation `ppaction`.
 *   - `{ kind: 'customShow', id, returnToShow }` — a named custom show.
 *
 * For `kind: 'slide'`, the matching slide is resolved by part name.
 * Returns `null` for unknown `ppaction` strings.
 */
declare const getShapeClickAction: (shape: SlideShapeData) => ShapeClickAction | null;
/**
 * Sets (or clears) the click action on the shape. Side effects:
 *
 *   - For `kind: 'url'`, a `hyperlink` rel is added (or reused) on the
 *     slide's rels with `targetMode="External"`. `<a:hlinkClick r:id=…/>`
 *     points at it.
 *   - For `kind: 'slide'`, a `slide` rel is added pointing at the
 *     target slide's part. The `<a:hlinkClick>` carries
 *     `action="ppaction://hlinksldjump"`.
 *   - For preset navigations and `customShow`, no rel is allocated; just the
 *     `action` attribute carries the destination.
 *   - `null` removes any existing `<a:hlinkClick>`.
 *
 * Optional `range` applies the action to selected UTF-16 text offsets instead
 * of the whole shape, preserving links and formatting outside the range.
 * Optional `tooltip` sets the link description; omission clears an old description.
 *
 * The shape must be one of `shape | picture | connector | graphicFrame`.
 * Groups don't carry their own click action in our model.
 */
declare const setShapeClickAction: (shape: SlideShapeData, action: ShapeClickAction | null, options?: {
  range?: {
    start: number;
    end: number;
  };
  tooltip?: string;
}) => void;
//#endregion
//#region src/internal/chartml/types.d.ts
/**
 * Chart type tokens. The kind names the data shape; the 3-D and
 * pie-of-pie variants are modifiers on `ChartSpec` rather than kinds of
 * their own:
 *
 *   | kind       | element              | modifier → element                        |
 *   | ---------- | -------------------- | ----------------------------------------- |
 *   | `bar`      | `<c:barChart>`       | `view3D` → `<c:bar3DChart>`               |
 *   | `column`   | `<c:barChart>`       | `view3D` → `<c:bar3DChart>`               |
 *   | `line`     | `<c:lineChart>`      | `view3D` → `<c:line3DChart>`              |
 *   | `area`     | `<c:areaChart>`      | `view3D` → `<c:area3DChart>`              |
 *   | `pie`      | `<c:pieChart>`       | `view3D` → `<c:pie3DChart>`, `ofPie` → `<c:ofPieChart>` |
 *   | `doughnut` | `<c:doughnutChart>`  |                                           |
 *   | `scatter`  | `<c:scatterChart>`   |                                           |
 *   | `bubble`   | `<c:bubbleChart>`    | `bubble3D` shades the bubbles             |
 *   | `radar`    | `<c:radarChart>`     |                                           |
 *   | `stock`    | `<c:stockChart>`     |                                           |
 *   | `surface`  | `<c:surface3DChart>` | `surfaceContour` → `<c:surfaceChart>`     |
 */
type ChartKind = 'bar' | 'column' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'radar' | 'bubble' | 'stock' | 'surface';
/** One labelled series of numeric values. */
interface ChartSeries {
  /** Series label rendered in the legend. */
  readonly name: string;
  /**
   * Numeric values, one per category. `null` slots become empty cells in
   * the embedded workbook (the reference desktop app draws them as a gap). Lengths
   * shorter than the category count are right-padded with `null`.
   *
   * For `scatter` / `bubble` kinds there are no categories: `values`
   * holds the series' y-channel (`<c:yVal>`), paired positionally with
   * `xValues`. For `radar` it behaves like `line` (values per category).
   *
   * A `stock` chart reads its series by position, not by name: three
   * series are high, low, close; four are open, high, low, close
   * (CT_StockChart allows exactly 3 or 4). A `surface` chart is a grid:
   * each series is one row of depth, each category one column.
   */
  readonly values: ReadonlyArray<number | null>;
  /**
   * X-channel values for `scatter` / `bubble` series (`<c:xVal>`), paired
   * positionally with `values`. Absent for every other kind, where
   * `values` is plotted against the shared `categories`.
   */
  readonly xValues?: ReadonlyArray<number | null>;
  /**
   * Per-point bubble sizes for `bubble` series (`<c:bubbleSize>`), paired
   * positionally with `xValues` / `values`. The renderer scales each
   * bubble's area (or width, per `ChartSpec.bubbleSizeRepresents`) to
   * this magnitude. Absent for non-bubble kinds.
   */
  readonly bubbleSizes?: ReadonlyArray<number | null>;
  /** Optional `#RRGGBB` fill override. Defaults to the theme's accent palette. */
  readonly color?: string;
  /**
   * Per-series chart-kind override for combo charts (e.g. a `line`
   * series overlaid on a `column` chart). Series sharing an effective
   * kind are grouped into one plot group (`<c:barChart>` /
   * `<c:lineChart>` / `<c:areaChart>`); the groups share the category
   * axis. Only the category kinds can mix — `bar` / `column` / `line` /
   * `area`. Absent = plotted with the chart's own `kind`.
   */
  readonly chartKind?: 'bar' | 'column' | 'line' | 'area';
  /**
   * Plot this series against the secondary value axis — the right-hand
   * axis the reference desktop app shows for combo charts with mixed units
   * (`<c:valAx>` pair with `axPos="r"` plus a deleted companion
   * `<c:catAx>`). The builder emits the secondary axis pair on demand.
   * Only meaningful for the category kinds; rejected for pie / doughnut.
   * At least one series must stay on the primary axis — a chart whose
   * every series is secondary is rejected at authoring time.
   */
  readonly secondaryAxis?: boolean;
  /**
   * Optional line stroke width in EMU (`<c:ser><c:spPr><a:ln w="…"/>`).
   * Only meaningful for line / area / scatter series. Default falls
   * back to the renderer's own pick.
   */
  readonly lineWidthEmu?: number;
  /**
   * Optional stroke color of the series outline / line
   * (`<c:ser><c:spPr><a:ln><a:solidFill>`), as an sRGB hex. Defaults to
   * `color`; on a pie / doughnut series it is the border between slices.
   */
  readonly lineColor?: string;
  /**
   * Optional line dash style token (`<c:ser><c:spPr><a:ln><a:prstDash
   * val="…"/>`), e.g. `'dash'`, `'dot'`, `'sysDash'`. Only
   * meaningful for line / area / scatter series.
   */
  readonly lineDash?: string;
  /**
   * Optional marker style for line / scatter series (`<c:ser><c:marker>`).
   * `none` hides the markers; the others render the matching glyph at
   * each data point. Defaults to `auto`, which renderers map to a small
   * filled circle.
   */
  readonly markerSymbol?: 'none' | 'auto' | 'circle' | 'square' | 'diamond' | 'triangle' | 'star' | 'x' | 'plus' | 'dash' | 'dot' | 'picture';
  /**
   * Marker size in points (`<c:marker><c:size val="N"/>`). The reference desktop app
   * default ~5. Only meaningful when `markerSymbol` isn't `none`.
   */
  readonly markerSizePt?: number;
  /**
   * Optional marker fill + outline color (`<c:marker><c:spPr>`), as an sRGB
   * hex. Defaults to `color`. The builder always writes it for line /
   * scatter / radar series, because the reference desktop app paints a marker without
   * `<c:spPr>` in the theme's automatic color instead of the series color.
   */
  readonly markerColor?: string;
  /**
   * Optional marker outline color (`<c:marker><c:spPr><a:ln>`), as an sRGB
   * hex. Defaults to `markerColor`, then `color`.
   */
  readonly markerLineColor?: string;
  /**
   * Invert the series color for negative values (bar / column charts).
   * Mirrors `<c:ser><c:invertIfNegative val="1"/>`. Renderers typically
   * paint the negative bars in the inverted shade of the series color.
   */
  readonly invertIfNegative?: boolean;
  /**
   * Optional per-data-point color overrides, indexed by point index
   * (`<c:dPt><c:idx val="N"/><c:spPr><a:solidFill>…`). Sparse — only
   * the indices that author an override appear. Pie / doughnut decks
   * almost always emit one of these per slice to break out of the
   * single-series-color default.
   */
  readonly pointColors?: ReadonlyArray<string | null>;
  /**
   * Optional per-data-point pie/doughnut slice explosion percentages
   * (`<c:dPt><c:explosion val="N"/>`). Sparse — only the indices that
   * author an explosion appear. The value is the radial offset as a
   * percentage of the slice radius (`25` ≈ a quarter-radius pull-out).
   */
  readonly pointExplosions?: ReadonlyArray<number | null>;
  /**
   * Optional per-data-point data-label overrides
   * (`<c:dLbls><c:dLbl><c:idx val="N"/>…`). Sparse — `null` slots fall back
   * to the series-level `dataLabels`. The reference desktop app draws the per-point element
   * over the series defaults, which is how pie / doughnut exporters give each
   * slice its own label content and font.
   */
  readonly pointDataLabels?: ReadonlyArray<ChartDataLabels | null>;
  /**
   * Line-smoothing toggle (`<c:smooth val="1"/>`) — only meaningful for
   * line / scatter / area series. When `true`, the renderer interpolates
   * a smooth curve through the data points instead of straight segments.
   */
  readonly smooth?: boolean;
  /**
   * Optional trendline overlay. ECMA-376 §21.2.2.211 allows several
   * regression types; we surface the most common subset. The line is
   * painted on top of the series in the renderer.
   */
  readonly trendline?: ChartTrendline;
  /**
   * Per-series data-label toggle overrides. Same shape as the chart-level
   * `ChartSpec.dataLabels`; overrides win when present. Read from
   * `<c:ser><c:dLbls>`.
   */
  readonly dataLabels?: ChartDataLabels;
  /**
   * Opacity of the series fill, `0` (transparent) to `1` (opaque) —
   * `<a:srgbClr><a:alpha val="…"/>` on the series' `<a:solidFill>`. Applies
   * to `color` and to every `pointColors` override. Overlapping area,
   * radar and bubble series are the usual reason to set it.
   */
  readonly fillOpacity?: number;
  /**
   * Error bars along the value direction (`<c:errBars>`; `<c:errDir
   * val="y"/>` on scatter / bubble). Valid on bar / column / line / area /
   * scatter / bubble series; rejected elsewhere, where CT_*Ser has no
   * `errBars` child.
   */
  readonly errorBars?: ChartErrorBars;
  /**
   * Error bars along the x direction (`<c:errDir val="x"/>`). Only scatter
   * and bubble series have an x channel to attach them to.
   */
  readonly xErrorBars?: ChartErrorBars;
}
/**
 * How far an error bar reaches (`<c:errValType>` plus its operand).
 * `stdErr` takes no operand; `cust` carries per-point literal amounts.
 */
type ChartErrorBarAmount = {
  readonly type: 'fixedVal' | 'percentage' | 'stdDev';
  readonly value: number;
} | {
  readonly type: 'stdErr';
} | {
  readonly type: 'cust'; /** Per-point positive amounts (`<c:plus><c:numLit>`). */
  readonly plus?: ReadonlyArray<number | null>; /** Per-point negative amounts (`<c:minus><c:numLit>`). */
  readonly minus?: ReadonlyArray<number | null>;
};
/** Error bars of one series in one direction (ECMA-376 §21.2.2.55). */
interface ChartErrorBars {
  /** Which side of the point gets a bar (`<c:errBarType>`). */
  readonly barType: 'both' | 'plus' | 'minus';
  readonly amount: ChartErrorBarAmount;
  /** Drop the T-shaped end caps (`<c:noEndCap val="1"/>`). */
  readonly noEndCap?: boolean;
  /** Bar stroke color as `#RRGGBB`. */
  readonly color?: string;
  /** Bar stroke width in EMU. */
  readonly lineWidthEmu?: number;
}
/** A single trendline overlay for a series. */
interface ChartTrendline {
  /**
   * Optional custom label for the trendline (`<c:trendline><c:name>…`).
   * Defaults to the reference desktop app's auto-generated label
   * (e.g. "Linear (X)" / "MA(5) (X)") when omitted.
   */
  readonly name?: string;
  /** Regression type — linear / exp / log / poly / power / movingAvg. */
  readonly type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
  /**
   * Forward extension (`<c:forward val="N"/>`) — the trendline runs N
   * data-point periods past the last point. Used to project future
   * values from the regression line.
   */
  readonly forward?: number;
  /**
   * Backward extension (`<c:backward val="N"/>`) — the trendline runs N
   * data-point periods before the first point.
   */
  readonly backward?: number;
  /** Optional moving-average period (only meaningful for type='movingAvg'). */
  readonly period?: number;
  /** Polynomial order (only meaningful for type='poly'). */
  readonly order?: number;
  /** Override stroke color; defaults to the series color. */
  readonly color?: string;
  /**
   * Show the regression equation next to the trendline
   * (`<c:dispEq val="1"/>`). Defaults to `false`.
   */
  readonly displayEquation?: boolean;
  /**
   * Show the R² coefficient next to the trendline
   * (`<c:dispRSqr val="1"/>`). Defaults to `false`.
   */
  readonly displayRSquared?: boolean;
}
/**
 * Per-series data-label toggles read from `<c:dLbls>` (ECMA-376
 * §21.2.2.55). All four toggles default to `false` — renderers paint
 * labels only when the corresponding flag is `true`.
 */
/**
 * `<c:dLblPos val="…"/>` — where the data label sits relative to its
 * data point. Per ECMA-376 §21.2.2.51 the token universe varies by
 * chart kind:
 *
 *   - bar / column: `'ctr' | 'inEnd' | 'outEnd' | 'inBase'`
 *   - line / area:  `'ctr' | 't' | 'b' | 'l' | 'r'`
 *   - pie / doughnut: `'ctr' | 'inEnd' | 'outEnd' | 'bestFit'`
 *
 * The union below covers every token; renderers ignore tokens that
 * don't apply to the chart's kind.
 */
type ChartDataLabelPosition = 'ctr' | 'inEnd' | 'outEnd' | 'inBase' | 't' | 'b' | 'l' | 'r' | 'bestFit';
interface ChartDataLabels {
  /** Numeric value of each data point. */
  readonly showValue: boolean;
  /** Category label of each data point. */
  readonly showCategory: boolean;
  /** Series name on each data point. */
  readonly showSeriesName: boolean;
  /** Percentage of total (for pie / doughnut). */
  readonly showPercent: boolean;
  /**
   * Number-format code from `<c:dLbls><c:numFmt formatCode="…"/>`. When
   * set, value labels are projected through this spreadsheet-style format
   * (same subset the value axis honors: `"0%"`, `"#,##0"`, `"$#,##0"`,
   * `"0.00"`). Independent of `ChartAxisScaling.numberFormat`.
   */
  readonly numberFormat?: string;
  /**
   * Position of the label relative to its data point — from
   * `<c:dLbls><c:dLblPos val="…"/>`.
   */
  readonly position?: ChartDataLabelPosition;
  /**
   * Separator between concatenated label parts when more than one of
   * `showValue` / `showCategory` / `showSeriesName` / `showPercent` is
   * set. Read from `<c:dLbls><c:separator>…</c:separator>` (a
   * leaf-text-only element). Common values: `" "` (default), `", "`,
   * `"\n"`, `"; "`.
   */
  readonly separator?: string;
  /**
   * Default-run text style for the labels — projected from
   * `<c:dLbls><c:txPr>…<a:defRPr/>…</c:txPr>`. Applies to whichever
   * label parts are turned on by the `show*` toggles. Independent of
   * the chart-title style and the axis-label style.
   */
  readonly textStyle?: ChartTextStyle;
  /**
   * Leader lines between a label and its data point
   * (`<c:dLbls><c:showLeaderLines val="…"/>`). Omitted = the
   * application default (shown). Series- and chart-level only: a
   * per-point `<c:dLbl>` has no such element, so `pointDataLabels`
   * ignores it on write and never reads it.
   */
  readonly showLeaderLines?: boolean;
  /** Bubble size of each data point (`<c:showBubbleSize>`); bubble charts only. */
  readonly showBubbleSize?: boolean;
  /** Legend key swatch next to each label (`<c:showLegendKey>`). */
  readonly showLegendKey?: boolean;
  /** Label background fill as `#RRGGBB` (`<c:spPr><a:solidFill>`). */
  readonly fillColor?: string;
  /**
   * Literal label text replacing the generated one (`<c:dLbl><c:tx>
   * <c:rich>`) — how a scatter point gets a name. Per-point only: the
   * series- and chart-level `<c:dLbls>` have no `<c:tx>`, so it is ignored
   * there on write and never read.
   */
  readonly text?: string;
}
/**
 * Authored value-axis scaling (`<c:valAx><c:scaling><c:min/>/<c:max/></c:scaling>`)
 * plus tick-spacing hints. When omitted, renderers compute the range
 * and tick spacing from the series values.
 */
interface ChartAxisScaling {
  readonly min?: number;
  readonly max?: number;
  /** Major tick spacing (`<c:majorUnit val="N"/>`). */
  readonly majorUnit?: number;
  /** Minor tick spacing (`<c:minorUnit val="N"/>`). */
  readonly minorUnit?: number;
  /**
   * Number-format code from `<c:numFmt formatCode="…"/>`. Common
   * values: `"0%"`, `"0.0%"`, `"#,##0"`, `"$#,##0"`, `"yyyy-mm-dd"`.
   * Renderers project a subset of spreadsheet number formats to label text.
   */
  readonly numberFormat?: string;
  /**
   * Logarithmic base for the axis from `<c:scaling><c:logBase val="N"/>`.
   * Typical values are `2`, `10`, `Math.E`. When set, the renderer
   * projects values through `Math.log(v) / Math.log(logBase)` before
   * mapping to plot coordinates. Linear when omitted.
   */
  readonly logBase?: number;
  /**
   * Authored display-units scale from `<c:dispUnits><c:builtInUnit val="…"/>`.
   * Token values map to divisors:
   *
   *   - `'hundreds'` (100), `'thousands'` (1e3), `'tenThousands'` (1e4),
   *     `'hundredThousands'` (1e5), `'millions'` (1e6),
   *     `'tenMillions'` (1e7), `'hundredMillions'` (1e8),
   *     `'billions'` (1e9), `'trillions'` (1e12)
   *
   * Renderers divide axis labels by the divisor and may append a unit
   * suffix ("K", "M", "B") for readability.
   */
  readonly displayUnits?: 'hundreds' | 'thousands' | 'tenThousands' | 'hundredThousands' | 'millions' | 'tenMillions' | 'hundredMillions' | 'billions' | 'trillions';
  /**
   * Show the display-units caption next to the axis ("Thousands",
   * "Millions", …) — an empty `<c:dispUnitsLbl/>` under `<c:dispUnits>`.
   * Only meaningful together with `displayUnits`.
   */
  readonly displayUnitsLabel?: boolean;
}
/**
 * 3-D view of a chart (`<c:view3D>`, ECMA-376 §21.2.2.228). Every field is
 * optional; an empty object still selects the 3-D plot-group element and
 * leaves the camera to the application's defaults.
 */
interface ChartView3D {
  /** Elevation in degrees, -90..90 (`<c:rotX>`). */
  readonly rotX?: number;
  /** Rotation about the vertical axis in degrees, 0..360 (`<c:rotY>`). */
  readonly rotY?: number;
  /**
   * Right-angle axes (`<c:rAngAx>`): an oblique projection that ignores
   * `perspective`. The reference desktop app's default for bar / column / line / area.
   */
  readonly rightAngleAxes?: boolean;
  /** Field of view, 0..240 (`<c:perspective>`); the reference desktop app's default is 30. */
  readonly perspective?: number;
  /** Depth as a percent of the chart width, 20..2000 (`<c:depthPercent>`). */
  readonly depthPercent?: number;
  /** Height as a percent of the chart width, 5..500 (`<c:hPercent>`). */
  readonly heightPercent?: number;
}
/**
 * Pie-of-pie / bar-of-pie settings (`<c:ofPieChart>`, ECMA-376
 * §21.2.2.126): the trailing points of a pie move into a second plot.
 */
interface ChartOfPie {
  /** Shape of the second plot (`<c:ofPieType>`). */
  readonly type: 'pie' | 'bar';
  /**
   * What decides which points move (`<c:splitType>`): `pos` — the last
   * `splitPos` points; `val` — points below `splitPos`; `percent` — points
   * below `splitPos` percent of the total; `cust` — the `customSplit`
   * indices; `auto` — the application's choice.
   */
  readonly splitType?: 'auto' | 'cust' | 'percent' | 'pos' | 'val';
  /** Threshold for `splitType` `pos` / `val` / `percent` (`<c:splitPos>`). */
  readonly splitPos?: number;
  /** Point indices in the second plot for `splitType: 'cust'` (`<c:custSplit>`). */
  readonly customSplit?: ReadonlyArray<number>;
  /** Second plot size as a percent of the first, 5..200 (`<c:secondPieSize>`). */
  readonly secondPieSizePct?: number;
  /** Gap between the two plots, 0..500 percent (`<c:gapWidth>`). */
  readonly gapWidthPct?: number;
  /** Connector lines between the two plots (`<c:serLines>`). */
  readonly seriesLines?: boolean;
}
/** Up / down bars between the first and last line series (`<c:upDownBars>`). */
interface ChartUpDownBars {
  /** Bar width gap, 0..500 percent (`<c:gapWidth>`); the reference desktop app's default is 150. */
  readonly gapWidthPct?: number;
  /** Fill of the rising bars as `#RRGGBB` (`<c:upBars>`). */
  readonly upColor?: string;
  /** Fill of the falling bars as `#RRGGBB` (`<c:downBars>`). */
  readonly downColor?: string;
}
/** Data table under the plot area (`<c:dTable>`, ECMA-376 §21.2.2.54). */
interface ChartDataTable {
  readonly showHorizontalBorder?: boolean;
  readonly showVerticalBorder?: boolean;
  readonly showOutline?: boolean;
  /** Legend keys next to the series names. */
  readonly showKeys?: boolean;
  readonly textStyle?: ChartTextStyle;
}
/** Calendar unit of a date axis (ST_TimeUnit). */
type ChartTimeUnit = 'days' | 'months' | 'years';
/**
 * Date-axis settings. Its presence turns the category axis into a
 * `<c:dateAx>`: `ChartSpec.categories` must then hold date serial numbers
 * (days since the workbook epoch, see `ChartSpec.date1904`) as strings,
 * and they are written as a numeric channel so the axis can space points
 * by date. Pair it with `categoryAxisNumberFormat` (e.g. `"yyyy-mm"`).
 */
interface ChartDateAxis {
  readonly baseTimeUnit?: ChartTimeUnit;
  readonly majorUnit?: number;
  readonly majorTimeUnit?: ChartTimeUnit;
  readonly minorUnit?: number;
  readonly minorTimeUnit?: ChartTimeUnit;
}
/**
 * Series (depth) axis of a 3-D chart (`<c:serAx>`). The builder emits the
 * axis whenever the plot-group element requires it; this object only
 * formats it.
 */
interface ChartSeriesAxis {
  readonly hidden?: boolean;
  readonly title?: string;
  readonly titleStyle?: ChartTextStyle;
  readonly labelStyle?: ChartTextStyle;
  readonly majorGridlines?: boolean;
  readonly orientation?: 'minMax' | 'maxMin';
  readonly tickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  readonly tickLabelSkip?: number;
  readonly lineColor?: string;
}
/**
 * Manual position of a chart element, as fractions (0..1) of the chart
 * space: `x` / `y` measured from its top-left corner, `w` / `h` as a
 * share of its size (`<c:layout><c:manualLayout>` with `xMode` / `yMode`
 * set to `edge`).
 */
interface ChartManualLayout {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  /**
   * Plot area only: whether the box bounds the plot alone (`inner`) or the
   * plot plus its tick labels and axis titles (`outer`, the default).
   */
  readonly target?: 'inner' | 'outer';
}
/**
 * Bar / column grouping per ECMA-376 §21.2.2.76 (`ST_BarGrouping`):
 *
 *   - `clustered` — bars within a category sit side-by-side (the default).
 *   - `stacked` — series values sit on top of each other; the y-axis spans
 *     0..max(sum of each category).
 *   - `percentStacked` — series values normalize to 100% per category.
 *   - `standard` — only meaningful for the 3D variants; renderers treat
 *     it as `clustered`.
 */
type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
/**
 * Authored text style for a chart label (title / axis title / etc.).
 * Read from the label's first `<a:rPr>` (and `<a:defRPr>` as fallback).
 * All fields are optional — absent fields mean "fall back to the
 * renderer's default for this label position."
 */
interface ChartTextStyle {
  /**
   * Font face applied to both the latin and east-asian typeface slots —
   * `<a:rPr><a:latin typeface="…"/><a:ea typeface="…"/>`. East-asian is
   * set alongside latin so a Japanese / CJK family (e.g. `'Yu Gothic'`)
   * renders the labels instead of the renderer's latin-only fallback.
   * Mirrors the latin/ea pairing the reference desktop app emits for CJK chart fonts.
   */
  readonly font?: string;
  /**
   * Complex-script font face — `<a:rPr><a:cs typeface="…"/>`. Written only
   * when set: `font` fills the latin and east-asian slots and never the
   * complex-script one, so a chart that mixes Arabic / Hebrew / Thai labels
   * has to name the face here as well.
   */
  readonly fontComplexScript?: string;
  /** Font size in points. From `<a:rPr sz="N"/>` where N is in 100ths of a pt. */
  readonly sizePt?: number;
  /** Bold flag from `<a:rPr b="1"/>`. */
  readonly bold?: boolean;
  /** Italic flag from `<a:rPr i="1"/>`. */
  readonly italic?: boolean;
  /** Fill color as `#RRGGBB` from `<a:rPr><a:solidFill><a:srgbClr/></a:solidFill>`. */
  readonly color?: string;
}
/**
 * Formatting for the secondary value axis (`<c:valAx>` with `axPos="r"`)
 * that `ChartSeries.secondaryAxis` series plot against. The fields mirror
 * the primary `valueAxis*` fields of `ChartSpec` without the prefix, limited
 * to what a right-hand axis needs; the axis position (`r`) and crossing
 * (`max`) stay fixed. Omit the whole object for the reference desktop app's defaults.
 */
interface ChartSecondaryValueAxis {
  readonly scaling?: ChartAxisScaling;
  readonly title?: string;
  readonly titleStyle?: ChartTextStyle;
  readonly labelStyle?: ChartTextStyle;
  readonly majorGridlines?: boolean;
  readonly majorGridlineColor?: string;
  /** Major gridline width in EMU, the mirror of `ChartSpec.valueAxisMajorGridlineWidthEmu`. */
  readonly majorGridlineWidthEmu?: number;
  readonly lineColor?: string;
  /** Axis line width in EMU, the mirror of `ChartSpec.valueAxisLineWidthEmu`. */
  readonly lineWidthEmu?: number;
  readonly majorTickMark?: 'in' | 'out' | 'cross' | 'none';
  readonly minorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Tick-label position, the mirror of `ChartSpec.valueAxisTickLabelPos`. */
  readonly tickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  readonly crossBetween?: 'between' | 'midCat';
}
/**
 * Every field a chart can carry, with no cross-field rules applied — the
 * shape `readChartSpec` returns. A deck authored elsewhere can combine
 * fields that this library would never write together, so the read type
 * stays permissive; `ChartSpec` is the closed, write-side domain.
 */
interface ReadChartSpec {
  readonly kind: ChartKind;
  /** Category labels along the x-axis (or pie slice labels for `kind: 'pie'`). */
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  /** Optional chart title rendered above the plot area. */
  readonly title?: string;
  /** Optional font / color overrides for the chart title. */
  readonly titleStyle?: ChartTextStyle;
  /** Optional chart-level data-label toggles. */
  readonly dataLabels?: ChartDataLabels;
  /** Optional value-axis scaling override (min / max). */
  readonly valueAxis?: ChartAxisScaling;
  /**
   * Formatting for the secondary value axis. Only meaningful when at least
   * one series sets `secondaryAxis: true`; ignored otherwise.
   */
  readonly secondaryValueAxis?: ChartSecondaryValueAxis;
  /**
   * Plot-area background fill — `<c:plotArea><c:spPr><a:solidFill>
   * <a:srgbClr val="…"/>`. `null` for no fill / unsupported fill kind.
   */
  readonly plotAreaFill?: string;
  /**
   * Plot-area outline color from `<c:plotArea><c:spPr><a:ln>
   * <a:solidFill><a:srgbClr/>`. `undefined` means no outline.
   */
  readonly plotAreaStrokeColor?: string;
  /**
   * Chart-area background fill — `<c:chartSpace><c:spPr><a:solidFill>
   * <a:srgbClr val="…"/>`. Renderers can use this as the outer card
   * color instead of the hard-coded white.
   */
  readonly chartAreaFill?: string;
  /**
   * Chart-area outline color from `<c:chartSpace><c:spPr><a:ln>
   * <a:solidFill><a:srgbClr/>`. Overrides the renderer's default
   * `#E5E7EB` card border.
   */
  readonly chartAreaStrokeColor?: string;
  /** Optional axis title text (`<c:catAx><c:title>` / `<c:valAx><c:title>`). */
  readonly categoryAxisTitle?: string;
  /** Authored font / color on the category-axis title (same `<a:rPr>` shape as `titleStyle`). */
  readonly categoryAxisTitleStyle?: ChartTextStyle;
  /**
   * Rotation of the category-axis title, in degrees clockwise. Maps
   * to `<c:catAx><c:title><c:tx><c:rich><a:bodyPr rot="N"/>` (the OOXML
   * value is in 60000ths of a degree, but the API surface uses plain
   * degrees). Omit to inherit the default (`0`).
   */
  readonly categoryAxisTitleRotationDeg?: number;
  /** Authored font / color on the category-axis *tick labels* — `<c:catAx><c:txPr>`. */
  readonly categoryAxisLabelStyle?: ChartTextStyle;
  /**
   * Authored rotation on the category-axis tick labels, in degrees. From
   * `<c:catAx><c:txPr><a:bodyPr rot="N"/>` where N is in 60000ths of a
   * degree. Positive values rotate clockwise (the reference desktop app's convention),
   * matching the SVG `transform=rotate()` sense.
   */
  readonly categoryAxisLabelRotationDeg?: number;
  readonly valueAxisTitle?: string;
  /** Authored font / color on the value-axis title. */
  readonly valueAxisTitleStyle?: ChartTextStyle;
  /**
   * Rotation of the value-axis title, in degrees clockwise. The reference desktop app
   * often emits `-90` (or `vert270`) so the title reads bottom-to-top
   * alongside the axis. Maps to `<c:valAx><c:title><c:tx><c:rich>
   * <a:bodyPr rot="N"/>` (60000ths of a degree on the wire).
   */
  readonly valueAxisTitleRotationDeg?: number;
  /** Authored font / color on the value-axis *tick labels* — `<c:valAx><c:txPr>`. */
  readonly valueAxisLabelStyle?: ChartTextStyle;
  /**
   * Authored rotation on the value-axis tick labels, in degrees. From
   * `<c:valAx><c:txPr><a:bodyPr rot="N"/>` (N in 60000ths of a degree).
   * Same sense as `categoryAxisLabelRotationDeg`.
   */
  readonly valueAxisLabelRotationDeg?: number;
  /** When `true`, value axis is hidden (`<c:valAx><c:delete val="1"/>`). */
  readonly valueAxisHidden?: boolean;
  /** When `true`, category axis is hidden (`<c:catAx><c:delete val="1"/>`). */
  readonly categoryAxisHidden?: boolean;
  /** When the value-axis emits `<c:majorGridlines/>` — its gridlines are visible. */
  readonly valueAxisMajorGridlines?: boolean;
  /**
   * Major-tick mark mode on the value axis (`<c:valAx><c:majorTickMark val="…"/>`):
   *
   *   - `'out'` — outside the plot edge (default)
   *   - `'in'` — inside the plot
   *   - `'cross'` — across the axis line
   *   - `'none'` — no tick marks
   */
  readonly valueAxisMajorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Major-tick mark mode on the category axis (`<c:catAx><c:majorTickMark>`). */
  readonly categoryAxisMajorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /** Minor-tick mark mode on the value axis (`<c:valAx><c:minorTickMark>`). */
  readonly valueAxisMinorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /**
   * Value-axis tick label position (`<c:valAx><c:tickLblPos val="…"/>`),
   * the same choices as `categoryAxisTickLabelPos`: `none` hides the
   * labels, `low` / `high` puts them at the plot's start / end, `nextTo`
   * (default) next to the axis. Primary axis only.
   */
  readonly valueAxisTickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  /** Minor-tick mark mode on the category axis (`<c:catAx><c:minorTickMark>`). */
  readonly categoryAxisMinorTickMark?: 'in' | 'out' | 'cross' | 'none';
  /**
   * Authored color on the value-axis major gridlines — `<c:valAx>
   * <c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="…"/>`.
   * Returned as `#RRGGBB`. `undefined` falls back to the renderer's
   * default (a light gray).
   */
  readonly valueAxisMajorGridlineColor?: string;
  /** Companion authored color on the value-axis minor gridlines. */
  readonly valueAxisMinorGridlineColor?: string;
  /** Authored color on the category-axis major gridlines (same shape as `valueAxisMajorGridlineColor`). */
  readonly categoryAxisMajorGridlineColor?: string;
  /** Authored color on the category-axis minor gridlines. */
  readonly categoryAxisMinorGridlineColor?: string;
  /**
   * Authored width of the value-axis major gridlines in EMU — the `w` of
   * the same `<a:ln>` as `valueAxisMajorGridlineColor` (12700 = 1 pt).
   * `undefined` leaves the width to the application (the reference desktop app: 0.75 pt).
   */
  readonly valueAxisMajorGridlineWidthEmu?: number;
  /** Companion width of the value-axis minor gridlines. */
  readonly valueAxisMinorGridlineWidthEmu?: number;
  /** Width of the category-axis major gridlines (same shape as `valueAxisMajorGridlineWidthEmu`). */
  readonly categoryAxisMajorGridlineWidthEmu?: number;
  /** Width of the category-axis minor gridlines. */
  readonly categoryAxisMinorGridlineWidthEmu?: number;
  /**
   * Authored color on the value-axis line itself — `<c:valAx><c:spPr>
   * <a:ln><a:solidFill><a:srgbClr val="…"/>`. Returned as `#RRGGBB`.
   * `undefined` falls back to the renderer's default axis stroke.
   */
  readonly valueAxisLineColor?: string;
  /** Authored color on the category-axis line (same shape as `valueAxisLineColor`). */
  readonly categoryAxisLineColor?: string;
  /**
   * Authored width of the value-axis line in EMU — the `w` of the same
   * `<a:ln>` as `valueAxisLineColor`. `undefined` leaves it to the application.
   */
  readonly valueAxisLineWidthEmu?: number;
  /** Width of the category-axis line (same shape as `valueAxisLineWidthEmu`). */
  readonly categoryAxisLineWidthEmu?: number;
  /** When the value-axis emits `<c:minorGridlines/>` — minor gridlines are visible. */
  readonly valueAxisMinorGridlines?: boolean;
  /**
   * Whether the category axis emits `<c:majorGridlines/>`. Sounds odd
   * for column charts (where category gridlines are vertical, between
   * categories) but bar charts put the cat axis on the vertical edge
   * and use these as horizontal guide lines per category band.
   */
  readonly categoryAxisMajorGridlines?: boolean;
  /** Companion `<c:minorGridlines/>` on the category axis. */
  readonly categoryAxisMinorGridlines?: boolean;
  /**
   * Category-axis tick label skip step (`<c:catAx><c:tickLblSkip val="N"/>`):
   * render every Nth category label. Commonly 2 / 5 / 10 on dense
   * time-series charts to keep labels from overlapping.
   */
  readonly categoryAxisTickLabelSkip?: number;
  /**
   * Category-axis tick *mark* skip (`<c:catAx><c:tickMarkSkip val="N"/>`):
   * draw every Nth tick mark independently of the label-skip stride.
   * Useful when you want fewer label collisions but the same dense
   * tick lattice; defaults to 1 (every tick).
   */
  readonly categoryAxisTickMarkSkip?: number;
  /**
   * Category-axis tick label position (`<c:catAx><c:tickLblPos val="…"/>`):
   *   - `none` hides labels but keeps the axis line
   *   - `low` / `high` puts them at the start / end (rare)
   *   - `nextTo` (default) is the standard position next to the axis
   */
  readonly categoryAxisTickLabelPos?: 'none' | 'low' | 'high' | 'nextTo';
  /**
   * Distance from the axis line to the labels, expressed as a percent of
   * the chart text size — `<c:catAx><c:lblOffset val="N"/>` where N is
   * 0..1000 (default 100). Larger values push category labels further
   * from the axis. Per ECMA-376 §21.2.2.94.
   */
  readonly categoryAxisLabelOffset?: number;
  /**
   * Multi-line category-label alignment relative to the tick mark —
   * `<c:catAx><c:lblAlgn val="ctr|l|r"/>`. The reference desktop app defaults to
   * `ctr` when omitted; the authored value wins.
   */
  readonly categoryAxisLabelAlign?: 'ctr' | 'l' | 'r';
  /**
   * Number-format code for the category-axis tick labels —
   * `<c:catAx><c:numFmt formatCode="…"/>`. Most useful on date-style
   * categories (`"mm/dd/yyyy"`, `"mmm"`, etc.) but accepts any spreadsheet
   * format string. Independent of `valueAxis.numberFormat` (which
   * targets the value axis).
   */
  readonly categoryAxisNumberFormat?: string;
  /**
   * Toggle multi-level (hierarchical) category labels —
   * `<c:catAx><c:noMultiLvlLbl val="0|1"/>`. The reference desktop app defaults to
   * `0` (multi-level labels stack). Set to `true` to flatten
   * hierarchical categories into a single row.
   */
  readonly categoryAxisNoMultiLevelLabel?: boolean;
  /**
   * Category-axis order — `'minMax'` (the data's natural order) or
   * `'maxMin'` (reversed). For bar charts the reference desktop app typically emits
   * `maxMin` so the first category sits at the top instead of the
   * bottom; honour the authored value when present.
   */
  readonly categoryAxisOrientation?: 'minMax' | 'maxMin';
  /** Same for the value axis. */
  readonly valueAxisOrientation?: 'minMax' | 'maxMin';
  /**
   * Where the category axis crosses the value axis. Either an enum
   * keyword (`<c:valAx><c:crosses val="autoZero|min|max"/>`) or a
   * specific numeric value (`<c:valAx><c:crossesAt val="N"/>`). The two
   * forms are mutually exclusive — the reference desktop app emits one or the other.
   * Default is `autoZero` (the category axis sits at value 0 if the
   * range straddles zero, otherwise at the closer extreme).
   */
  readonly valueAxisCrosses?: 'autoZero' | 'min' | 'max' | {
    at: number;
  };
  /**
   * Whether the value axis crosses the category axis *between* tick
   * marks (the default for bar / column / area) or *at* each tick mark
   * (the default for line / scatter). Maps to `<c:valAx>
   * <c:crossBetween val="between|midCat"/>`. The reference desktop app emits this when
   * the chart kind makes the default value non-obvious — surface it
   * here so the round-trip preserves the authored intent.
   */
  readonly valueAxisCrossBetween?: 'between' | 'midCat';
  /**
   * When `false`, plot data from hidden cells in the embedded workbook
   * — maps to `<c:plotVisOnly val="0"/>`. The reference desktop app's default is
   * `true` (only plot visible cells), so omitting this field emits
   * `val="1"` to stay round-trip-safe with files the reference desktop app authored.
   */
  readonly plotVisibleCellsOnly?: boolean;
  /**
   * Renders the chart area with rounded corners
   * (`<c:chartSpace><c:roundedCorners val="1"/>`). The reference desktop app's default
   * is `false`; surface only when explicitly `true` so the round-trip
   * doesn't add a redundant `false`.
   */
  readonly roundedCorners?: boolean;
  /**
   * The reference desktop app's built-in chart-style preset
   * (`<c:chartSpace><c:style val="N"/>`), 1–48. Encodes a curated combo
   * of theme accent colors, gradients, effects, and font sizes that
   * the reference desktop app applies when the user picks a chart style from the
   * "Chart Styles" gallery. Surface for round-trip parity; renderers in
   * @office-kit/pptx don't (yet) interpret it.
   */
  readonly chartStyle?: number;
  /**
   * Language code for the chart's number / date formatters
   * (`<c:chartSpace><c:lang val="…"/>`). The reference desktop app emits the user's
   * Authoring app UI language (e.g. `'en-US'`, `'ja-JP'`). Carried for
   * round-trip parity; renderers in @office-kit/pptx don't act on it yet.
   */
  readonly language?: string;
  /**
   * `<c:chartSpace><c:date1904 val="…"/>` — the spreadsheet date-system flag.
   * `false` (the default) uses the 1900-epoch; `true` uses 1904.
   * Surface for parity; renderers don't act on it yet.
   */
  readonly date1904?: boolean;
  /** Bar / column / area grouping mode. Absent for line / pie. */
  readonly grouping?: ChartGrouping;
  /**
   * Drop lines (`<c:dropLines>`) — vertical guide lines from each data
   * point down to the value axis. Common on line / area charts.
   */
  readonly dropLines?: boolean;
  /**
   * High-low lines (`<c:hiLowLines>`) — vertical lines spanning the
   * highest and lowest series value at each category. Used by stock
   * charts and side-by-side line charts.
   */
  readonly hiLowLines?: boolean;
  /**
   * Whether a line chart draws point markers (`<c:lineChart><c:marker
   * val="1"/>`). `true` is the reference desktop app's "Line with Markers" subtype;
   * absent / `false` is the plain "Line" subtype (no markers). Only
   * meaningful for `kind: 'line'`.
   */
  readonly lineMarkers?: boolean;
  /**
   * Gap between adjacent bar groups in `<c:gapWidth val="N"/>` units
   * (0..500, percent of bar width). Default 150 (= 1.5×) in the reference desktop app.
   */
  readonly gapWidthPct?: number;
  /**
   * Overlap of adjacent bars within a category in `<c:overlap val="N"/>`
   * percent (-100..100). Negative pulls bars apart, positive overlaps.
   * Defaults to 0 (clustered) / 100 (stacked).
   */
  readonly overlapPct?: number;
  /**
   * Optional legend configuration. `position` mirrors
   * `<c:legend><c:legendPos val="…"/>` — `'r'` (right) is the default,
   * `'t'`, `'b'`, `'l'`, `'tr'` (top-right) the other tokens. `null`
   * for `position` means no legend at all (the chart explicitly hides
   * it). `undefined` overall means the chart didn't author a legend
   * element — renderers should fall back to their own default.
   */
  readonly legend?: {
    position: 'r' | 't' | 'b' | 'l' | 'tr' | null; /** When `true`, legend overlays the plot area instead of taking a strip. */
    readonly overlay?: boolean; /** Manual position of the legend box; `target` does not apply. */
    readonly layout?: ChartManualLayout;
    /**
     * Authored font / color on the legend's text. From `<c:legend><c:txPr>`'s
     * first `<a:p><a:pPr><a:defRPr>` (or `<a:r><a:rPr>` as fallback). Same
     * shape as `titleStyle`.
     */
    readonly textStyle?: ChartTextStyle;
    /**
     * Series indices that the legend hides via `<c:legend><c:legendEntry>
     * <c:idx val="N"/><c:delete val="1"/></c:legendEntry>`. Renderers
     * filter them from the legend list while still plotting them.
     * Common use: keep a trendline series in the data but drop its
     * legend entry.
     */
    readonly hiddenIndices?: ReadonlyArray<number>;
  };
  /** When `true`, the chart title overlays the plot area instead of taking a strip. */
  readonly titleOverlay?: boolean;
  /**
   * `<c:barChart><c:varyColors val="1"/>` etc. — when `true` and the
   * chart has a single series, each data point gets a distinct color
   * from the palette. Pie / doughnut already vary colors implicitly;
   * this flag is most useful for single-series bar / column.
   */
  readonly varyColors?: boolean;
  /**
   * `<c:dispBlanksAs val="…"/>` — how line / area renderers should
   * treat `null` values in series:
   *
   *   - `'gap'`  — leave a gap (the default)
   *   - `'zero'` — substitute zero
   *   - `'span'` — connect the surrounding points across the gap
   */
  readonly dispBlanksAs?: 'gap' | 'zero' | 'span';
  /**
   * Pie / doughnut: angle (in degrees, 0–360) at which the first slice
   * starts, measured clockwise from 12 o'clock. Mirrors
   * `<c:firstSliceAng val="…"/>`. Default 0 = start at the top.
   */
  readonly firstSliceAngleDeg?: number;
  /**
   * Doughnut hole size as a percent of the outer radius (10..90).
   * Mirrors `<c:holeSize val="…"/>`. Default 50.
   */
  readonly holeSizePct?: number;
  /**
   * Scatter sub-type from `<c:scatterChart><c:scatterStyle val="…"/>`
   * (ECMA-376 ST_ScatterStyle). Governs whether the renderer connects
   * points with straight / smoothed lines and whether it draws markers:
   *
   *   - `'marker'` (and the renderer's default when absent) — markers only
   *   - `'line'` / `'smooth'` — connecting line only
   *   - `'lineMarker'` / `'smoothMarker'` — line + markers
   *   - `'none'` — neither
   */
  readonly scatterStyle?: 'none' | 'line' | 'lineMarker' | 'marker' | 'smooth' | 'smoothMarker';
  /**
   * Radar sub-type from `<c:radarChart><c:radarStyle val="…"/>`
   * (ECMA-376 ST_RadarStyle): `'standard'` (polyline only), `'marker'`
   * (polyline + point markers), or `'filled'` (the polygon is filled at
   * reduced opacity).
   */
  readonly radarStyle?: 'standard' | 'marker' | 'filled';
  /**
   * Bubble-size scale percentage from `<c:bubbleChart><c:bubbleScale
   * val="N"/>` (0..300, the reference desktop app's default 100). Scales every bubble's
   * rendered radius proportionally.
   */
  readonly bubbleScale?: number;
  /**
   * Whether a bubble's `<c:bubbleSize>` maps to the bubble's area or its
   * width — `<c:bubbleChart><c:sizeRepresents val="area|w"/>` (ECMA-376
   * ST_SizeRepresents; the `'w'` token is surfaced as `'width'`).
   * The reference desktop app's default is `'area'`, so radius scales with `sqrt(size)`;
   * `'width'` scales radius linearly with size.
   */
  readonly bubbleSizeRepresents?: 'area' | 'width';
  /**
   * Shade bubbles as spheres — `<c:bubble3D val="1"/>` on each series of the
   * bubble chart (the reference desktop app's "3-D Bubble" subtype). Unlike the
   * other 3-D variants it has no camera, so `view3D` does not apply.
   */
  readonly bubble3D?: boolean;
  /** Plot bubbles with a negative size (`<c:showNegBubbles>`). */
  readonly showNegativeBubbles?: boolean;
  /**
   * 3-D view. Its presence selects the 3-D plot-group element for `bar` /
   * `column` / `line` / `area` / `pie`; for `surface` it only positions the
   * camera. Rejected for every other kind, and for combo charts (a 3-D
   * group cannot share a plot area with another group in the reference desktop app).
   */
  readonly view3D?: ChartView3D;
  /** Shape of 3-D bars / columns (`<c:bar3DChart><c:shape>`). Requires `view3D`. */
  readonly bar3DShape?: 'box' | 'cone' | 'coneToMax' | 'cylinder' | 'pyramid' | 'pyramidToMax';
  /**
   * Depth gap between the series rows of a 3-D bar / column / line / area
   * chart, 0..500 percent (`<c:gapDepth>`). Requires `view3D`.
   */
  readonly gapDepthPct?: number;
  /** Series (depth) axis formatting for 3-D charts and `surface`. */
  readonly seriesAxis?: ChartSeriesAxis;
  /** Turns a `pie` into a pie-of-pie / bar-of-pie chart. */
  readonly ofPie?: ChartOfPie;
  /**
   * Surface charts: `true` draws the top-down contour plot
   * (`<c:surfaceChart>`), absent / `false` the 3-D surface
   * (`<c:surface3DChart>`).
   */
  readonly surfaceContour?: boolean;
  /** Surface charts: draw the mesh only, without color bands (`<c:wireframe>`). */
  readonly surfaceWireframe?: boolean;
  /**
   * Up / down bars. Valid on `line` charts (between the first and last
   * series) and on `stock` charts with an open series, where they are the
   * candlestick bodies.
   */
  readonly upDownBars?: ChartUpDownBars;
  /** Data table under the plot area. Rejected for the axis-less kinds. */
  readonly dataTable?: ChartDataTable;
  /** Turns the category axis into a date axis — see `ChartDateAxis`. */
  readonly categoryAxisDate?: ChartDateAxis;
  /**
   * Scaling of the horizontal axis where it is numeric: the x value axis
   * of `scatter` / `bubble` charts (every field applies) and a date axis
   * (`min` / `max` only, as date serials). Rejected on a text category
   * axis, which has no scale to set.
   */
  readonly categoryAxisScaling?: ChartAxisScaling;
  /**
   * Outer levels of a multi-level category axis (`<c:multiLvlStrRef>`),
   * ordered from the level next to `categories` outward. Each level has
   * one entry per category; an empty string continues the previous
   * group, so `['2024', '', '', '', '2025', '', '', '']` groups eight
   * quarters into two years.
   */
  readonly categoryGroupLevels?: ReadonlyArray<ReadonlyArray<string>>;
  /** Manual position of the plot area. */
  readonly plotAreaLayout?: ChartManualLayout;
  /**
   * Manual position of the chart title's top-left corner, as fractions of
   * the chart space. A title sizes itself to its text, so there is no
   * width / height to set.
   */
  readonly titleLayout?: {
    readonly x: number;
    readonly y: number;
  };
  /**
   * Hide the value-axis line while keeping its tick labels
   * (`<c:valAx><c:spPr><a:ln><a:noFill/>`). `valueAxisHidden` removes the
   * labels too; this does not.
   */
  readonly valueAxisLineHidden?: boolean;
  /** Hide the category-axis line while keeping its labels. */
  readonly categoryAxisLineHidden?: boolean;
}
//#endregion
//#region src/internal/chartml/chart-spec.d.ts
/** Field of a chart whose availability depends on the kind. */
type SpecField = Exclude<keyof ReadChartSpec, 'kind' | 'series'>;
/**
 * A kind's field set: `K` as declared on `ReadChartSpec`, and every other
 * field closed off. The `?: never` arms are what turn a field the kind
 * cannot draw into a compile error.
 */
type Fields<K extends SpecField> = Pick<ReadChartSpec, K> & { readonly [P in Exclude<SpecField, K>]?: never };
/** Per-variant field sets; the variant types below are derived from these. */
declare const VARIANT_FIELDS: {
  readonly combo: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "grouping", "gapWidthPct", "overlapPct", "dropLines", "hiLowLines", "lineMarkers", "upDownBars", "secondaryValueAxis"];
  readonly bar3D: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "grouping", "gapWidthPct", "gapDepthPct", "bar3DShape", "seriesAxis", "view3D"];
  readonly lineArea3D: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "grouping", "dropLines", "gapDepthPct", "seriesAxis", "view3D"];
  readonly pie: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "firstSliceAngleDeg"];
  readonly ofPie: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "ofPie"];
  readonly pie3D: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "view3D"];
  readonly doughnut: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "firstSliceAngleDeg", "holeSizePct"];
  readonly scatter: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisScaling", "scatterStyle"];
  readonly bubble: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisScaling", "bubbleScale", "bubbleSizeRepresents", "bubble3D", "showNegativeBubbles"];
  readonly radar: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "radarStyle"];
  readonly stock: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "dropLines", "hiLowLines"];
  readonly candlestick: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "dropLines", "hiLowLines", "upDownBars"];
  readonly surface: readonly ["categories", "title", "titleStyle", "titleOverlay", "titleLayout", "dataLabels", "legend", "plotAreaLayout", "plotAreaFill", "plotAreaStrokeColor", "chartAreaFill", "chartAreaStrokeColor", "plotVisibleCellsOnly", "dispBlanksAs", "date1904", "language", "roundedCorners", "chartStyle", "varyColors", "valueAxis", "valueAxisTitle", "valueAxisTitleStyle", "valueAxisTitleRotationDeg", "valueAxisLabelStyle", "valueAxisLabelRotationDeg", "valueAxisHidden", "valueAxisMajorGridlines", "valueAxisMinorGridlines", "valueAxisMajorGridlineColor", "valueAxisMinorGridlineColor", "valueAxisMajorGridlineWidthEmu", "valueAxisMinorGridlineWidthEmu", "valueAxisMajorTickMark", "valueAxisMinorTickMark", "valueAxisTickLabelPos", "valueAxisLineColor", "valueAxisLineWidthEmu", "valueAxisLineHidden", "valueAxisOrientation", "valueAxisCrosses", "valueAxisCrossBetween", "categoryAxisTitle", "categoryAxisTitleStyle", "categoryAxisTitleRotationDeg", "categoryAxisLabelStyle", "categoryAxisLabelRotationDeg", "categoryAxisHidden", "categoryAxisMajorGridlines", "categoryAxisMinorGridlines", "categoryAxisMajorGridlineColor", "categoryAxisMinorGridlineColor", "categoryAxisMajorGridlineWidthEmu", "categoryAxisMinorGridlineWidthEmu", "categoryAxisMajorTickMark", "categoryAxisMinorTickMark", "categoryAxisTickLabelPos", "categoryAxisLineColor", "categoryAxisLineWidthEmu", "categoryAxisLineHidden", "categoryAxisOrientation", "categoryAxisNumberFormat", "categoryAxisTickLabelSkip", "categoryAxisTickMarkSkip", "categoryAxisLabelOffset", "categoryAxisLabelAlign", "categoryAxisNoMultiLevelLabel", "categoryAxisDate", "categoryAxisScaling", "categoryGroupLevels", "dataTable", "surfaceContour", "surfaceWireframe", "view3D", "seriesAxis"];
};
/** Name of one arm of {@link ChartSpec}. */
type ChartVariant = keyof typeof VARIANT_FIELDS;
type VariantField<V extends ChartVariant> = (typeof VARIANT_FIELDS)[V][number];
/**
 * Series channels and combo placement are the only parts of a series that
 * depend on the chart kind; name, styling and per-point overrides are shared.
 */
type ChartSeriesCommon = Omit<ChartSeries, 'xValues' | 'bubbleSizes' | 'chartKind' | 'secondaryAxis'>;
/** A series plotted as one value channel against the categories. */
type PlainChartSeries = ChartSeriesCommon & {
  readonly xValues?: never;
  readonly bubbleSizes?: never;
  readonly chartKind?: never;
  readonly secondaryAxis?: never;
};
/**
 * A series on a bar / column / line / area chart. Those four share a plot
 * area, so a series can be drawn as another of them (`chartKind`) or moved
 * to the secondary value axis.
 */
type ComboChartSeries = ChartSeriesCommon & {
  readonly xValues?: never;
  readonly bubbleSizes?: never;
  readonly chartKind?: 'bar' | 'column' | 'line' | 'area';
  readonly secondaryAxis?: boolean;
};
/** A scatter series: an x channel paired with the y channel. */
type ScatterChartSeries = ChartSeriesCommon & {
  readonly xValues: ReadonlyArray<number | null>;
  readonly bubbleSizes?: never;
  readonly chartKind?: never;
  readonly secondaryAxis?: never;
};
/** A bubble series: x, y, and the size channel that scales each bubble. */
type BubbleChartSeries = ChartSeriesCommon & {
  readonly xValues: ReadonlyArray<number | null>;
  readonly bubbleSizes: ReadonlyArray<number | null>;
  readonly chartKind?: never;
  readonly secondaryAxis?: never;
};
/**
 * A category axis is either discrete — optionally grouped into outer
 * levels — or a date axis. Only a numeric axis has a scale, so
 * `categoryAxisScaling` rides with the date form.
 */
type CategoryAxisMode = {
  readonly categoryAxisDate: ChartDateAxis; /** Date-serial range of the axis; only `min` / `max` apply. */
  readonly categoryAxisScaling?: ChartAxisScaling;
  readonly categoryGroupLevels?: never;
} | {
  readonly categoryAxisDate?: never;
  readonly categoryAxisScaling?: never;
  readonly categoryGroupLevels?: ReadonlyArray<ReadonlyArray<string>>;
};
/**
 * A bar / column / line / area chart. Because the four share a plot area,
 * every series-level and group-level control of the family is available
 * whichever of them is the base kind — that is what makes a combo chart.
 */
type ComboChartSpec = Fields<VariantField<'combo'>> & CategoryAxisMode & {
  readonly kind: 'bar' | 'column' | 'line' | 'area';
  readonly series: ReadonlyArray<ComboChartSeries>;
};
/** A 3-D bar / column chart (`<c:bar3DChart>`). */
type Bar3DChartSpec = Fields<VariantField<'bar3D'>> & CategoryAxisMode & {
  readonly kind: 'bar' | 'column';
  readonly series: ReadonlyArray<PlainChartSeries>; /** Its presence is what selects the 3-D element; a combo cannot be 3-D. */
  readonly view3D: ChartView3D;
};
/**
 * A 3-D line / area chart. High-low lines, up/down bars and point markers
 * have no 3-D element to live in.
 */
type LineArea3DChartSpec = Fields<VariantField<'lineArea3D'>> & CategoryAxisMode & {
  readonly kind: 'line' | 'area';
  readonly series: ReadonlyArray<PlainChartSeries>;
  readonly view3D: ChartView3D;
};
/** A pie chart. One series, one slice per category. */
type PieChartSpec = Fields<VariantField<'pie'>> & {
  readonly kind: 'pie';
  readonly series: readonly [PlainChartSeries];
};
/** A pie-of-pie / bar-of-pie chart (`<c:ofPieChart>`). */
type OfPieChartSpec = Fields<VariantField<'ofPie'>> & {
  readonly kind: 'pie';
  readonly series: readonly [PlainChartSeries];
  readonly ofPie: ChartOfPie;
};
/** A 3-D pie (`<c:pie3DChart>`), which turns with the camera rather than `firstSliceAngleDeg`. */
type Pie3DChartSpec = Fields<VariantField<'pie3D'>> & {
  readonly kind: 'pie';
  readonly series: readonly [PlainChartSeries];
  readonly view3D: ChartView3D;
};
/** A doughnut chart. */
type DoughnutChartSpec = Fields<VariantField<'doughnut'>> & {
  readonly kind: 'doughnut';
  readonly series: readonly [PlainChartSeries];
};
/** A scatter chart: both axes are numeric, so there are no categories to group. */
type ScatterChartSpec = Fields<VariantField<'scatter'>> & {
  readonly kind: 'scatter';
  readonly series: ReadonlyArray<ScatterChartSeries>;
};
/** A bubble chart: a scatter chart whose third channel scales each marker. */
type BubbleChartSpec = Fields<VariantField<'bubble'>> & {
  readonly kind: 'bubble';
  readonly series: ReadonlyArray<BubbleChartSeries>;
};
/** A radar chart. */
type RadarChartSpec = Fields<VariantField<'radar'>> & CategoryAxisMode & {
  readonly kind: 'radar';
  readonly series: ReadonlyArray<PlainChartSeries>;
};
/** A stock chart plotted from high, low and close. */
type StockChartSpec = Fields<VariantField<'stock'>> & CategoryAxisMode & {
  readonly kind: 'stock';
  readonly series: readonly [PlainChartSeries, PlainChartSeries, PlainChartSeries];
};
/**
 * A candlestick chart: a stock chart with an open series, whose up / down
 * bars span open → close.
 */
type CandlestickChartSpec = Fields<VariantField<'candlestick'>> & CategoryAxisMode & {
  readonly kind: 'stock';
  readonly series: readonly [PlainChartSeries, PlainChartSeries, PlainChartSeries, PlainChartSeries];
};
/** A surface chart. Its three axes make `view3D` a camera, not a subtype switch. */
type SurfaceChartSpec = Fields<VariantField<'surface'>> & CategoryAxisMode & {
  readonly kind: 'surface';
  readonly series: ReadonlyArray<PlainChartSeries>;
};
/**
 * A chart this library can draw. Each kind carries only the fields its
 * OOXML element has, so a field the chart cannot show is a compile error
 * rather than a value that serializes into nothing — or into a deck
 * the reference desktop app offers to repair.
 *
 * `readChartSpec` returns the permissive `ReadChartSpec` instead, because a
 * deck authored elsewhere can combine fields no kind allows; narrow one back
 * with {@link isChartSpec}.
 */
type ChartSpec = ComboChartSpec | Bar3DChartSpec | LineArea3DChartSpec | PieChartSpec | OfPieChartSpec | Pie3DChartSpec | DoughnutChartSpec | ScatterChartSpec | BubbleChartSpec | RadarChartSpec | StockChartSpec | CandlestickChartSpec | SurfaceChartSpec;
/**
 * Whether a spec read back from a deck also inhabits the write-side
 * {@link ChartSpec}. A deck authored elsewhere can combine fields no kind
 * draws, so a read-modify-write pass has to narrow before it writes.
 */
declare const isChartSpec: (spec: ReadChartSpec) => spec is ChartSpec;
//#endregion
//#region src/api/fn/charts.d.ts
/**
 * Adds a native chart to the slide at `x`/`y` with size `w`×`h` (EMU),
 * built from `spec`, with an embedded workbook so the reference desktop app's "Edit data"
 * opens the chart's numbers. Returns the chart's graphic-frame shape.
 */
declare const addSlideChart: (slide: SlideData, opts: {
  spec: ChartSpec;
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  name?: string;
}) => SlideShapeData;
/**
 * A chart sitting on a slide. `shape` is the `<p:graphicFrame>`
 * wrapper; `spec` is the chart definition parsed from the linked
 * `/ppt/charts/chart{N}.xml` part. `null` `spec` means the chart uses
 * a kind we don't model (callers can fall through to pass-through).
 */
interface SlideChartData {
  readonly shape: SlideShapeData;
  readonly spec: ReadChartSpec | null;
}
/**
 * Replaces the chart definition on an existing graphic-frame chart
 * shape. Updates the inline `<c:strCache>` / `<c:numCache>` blocks so
 * the reference desktop app renders the new data without opening the embedded
 * workbook. The shape's geometry (position / size / rotation) is
 * preserved verbatim.
 *
 * The embedded xlsx is re-written too — it's what the "Edit data"
 * affordance opens. The previous workbook is replaced wholesale (no
 * attempt to preserve styles a user added in a spreadsheet app).
 *
 * Pass any `ChartSpec`, including a different `kind` from the
 * original; this acts as "change my column chart to a line chart with
 * fresh data."
 */
declare const setChartSpec: (chart: SlideChartData, spec: ChartSpec) => void;
/**
 * Returns every chart on the slide, with its `ChartSpec` parsed from
 * the linked chart part. Skips graphic frames that don't carry a
 * `<c:chart>` reference (e.g. tables, diagrams).
 */
/**
 * For a graphic-frame shape that wraps a chart, returns the parsed
 * `ChartSpec`. Returns `null` when the shape isn't a chart wrapper
 * or the chart uses a kind we don't model (e.g. surface, radar).
 *
 * Convenience over `getSlideCharts(...).find((c) => c.shape === shape)`
 * when the caller already has the shape in hand (e.g. iterating
 * `getSlideShapes`).
 */
/**
 * Convenience over `getShapeChartSpec(shape)?.kind ?? null`. Returns
 * the chart's `ChartKind` ('bar', 'line', 'pie', …) when the shape
 * is a chart wrapper, or `null` for non-charts and charts whose
 * kind isn't modeled yet.
 */
declare const getShapeChartKind: (shape: SlideShapeData) => ChartKind | null;
/**
 * Returns the categories axis labels of a chart shape, or `null`
 * if the shape isn't a chart wrapper or its kind isn't modeled.
 * Convenience over `getShapeChartSpec(shape)?.categories ?? null`.
 */
declare const getShapeChartCategories: (shape: SlideShapeData) => ReadonlyArray<string> | null;
/**
 * Returns the chart's series-name list (in spec order). `null`
 * when the shape isn't a chart wrapper or the kind isn't modeled.
 */
declare const getShapeChartSeriesNames: (shape: SlideShapeData) => ReadonlyArray<string> | null;
/**
 * Returns the values for the named series on a chart shape, or
 * `null` when the shape isn't a chart, the kind isn't modeled, or
 * no series matches `seriesName`.
 *
 * For `scatter` / `bubble` charts the returned array is the series'
 * y-channel (`<c:yVal>`); the paired x-channel and bubble sizes are on
 * the full spec via `getShapeChartSpec` (`series[].xValues` /
 * `series[].bubbleSizes`).
 */
declare const getShapeChartSeriesValues: (shape: SlideShapeData, seriesName: string) => ReadonlyArray<number | null> | null;
declare const getShapeChartSpec: (shape: SlideShapeData) => ReadChartSpec | null;
/**
 * Returns every chart on the slide that carries a series whose name
 * equals `seriesName` exactly. Useful for "find the revenue chart"
 * patterns where chart kind alone isn't unique. Skips charts whose
 * kind isn't modeled.
 */
declare const findChartsBySeriesName: (slide: SlideData, seriesName: string) => ReadonlyArray<SlideChartData>;
/**
 * Returns every chart on the slide that carries at least one
 * `<c:trendline>` on any of its series. Useful for deck-audit reports
 * — trendlines are easy to add and easy to forget, and authors often
 * want to inventory them before publishing. Skips charts whose kind
 * isn't modeled.
 */
declare const findChartsWithTrendlines: (slide: SlideData) => ReadonlyArray<SlideChartData>;
/**
 * Returns every chart on the slide whose `dataLabels` (chart-level)
 * or any series-level `dataLabels` enable at least one of `showValue`,
 * `showCategory`, `showSeriesName`, or `showPercent`. Useful for
 * "which charts show numbers next to the data?" audits — purely
 * presence-based; doesn't validate numberFormat or position. Charts
 * whose kind isn't modeled are skipped.
 */
declare const findChartsWithDataLabels: (slide: SlideData) => ReadonlyArray<SlideChartData>;
/**
 * Returns the first chart on the slide whose parsed `kind` matches
 * `kind` (e.g. `'bar'`, `'line'`, `'pie'`). Returns `null` when no
 * chart on the slide has that kind, or when every chart on the slide
 * uses a kind this version doesn't yet model.
 */
declare const findChartByKind: (slide: SlideData, kind: ChartKind) => SlideChartData | null;
/**
 * Histogram of chart kinds across the whole deck. Returns a frozen
 * `Record<ChartKind, number>` with every kind present (zeros for any
 * absent kind). Useful for deck-audit reports (e.g. "how many pie
 * charts does this presentation have?") without iterating every slide
 * by hand.
 *
 * Charts whose spec doesn't parse — typically a kind this version
 * doesn't yet model — are skipped silently, mirroring
 * `findChartByKind`'s null-spec handling.
 */
declare const getPresentationChartKindCounts: (pres: PresentationData) => Readonly<Record<ChartKind, number>>;
/**
 * Dense per-slide chart count array, 0-based by slide index. Counts
 * every chart returned by `getSlideCharts` regardless of whether its
 * `spec` parsed — useful for "which slides carry charts the renderer
 * doesn't yet model?" audits. Pair with
 * `getPresentationChartKindCounts` for kind-level totals.
 */
declare const getPresentationChartCountsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
declare const getSlideCharts: (slide: SlideData) => ReadonlyArray<SlideChartData>;
//#endregion
//#region src/api/fn/tables.d.ts
/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping `<a:tbl>` —
 * i.e. a table. Sharper than `getShapeKind(shape) === 'graphicFrame'`,
 * which also matches charts and SmartArt frames.
 */
declare const isTableShape: (shape: SlideShapeData) => boolean;
/**
 * Returns every table graphic-frame shape on the slide, in document
 * order. Convenience over `getSlideShapes(slide).filter(isTableShape)`
 * — pairs `getSlideCharts(slide)` for cases where the caller wants
 * just the tables.
 */
declare const getSlideTables: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/**
 * Dense per-slide table count array. Counts `isTableShape` matches per
 * slide. Rounds out the density-array family alongside the chart,
 * shape, text, and comment counters.
 */
declare const getPresentationTableCountsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping a chart
 * reference (`<c:chart>`). Charts, tables, and SmartArt all share
 * the graphic-frame kind; this predicate filters down to charts only.
 */
declare const isChartShape: (shape: SlideShapeData) => boolean;
/**
 * Returns a 2D array of cell handles for the given table shape, in
 * row-major order. Throws if the shape isn't a table graphic frame.
 */
declare const getTableCells: (table: SlideShapeData) => ReadonlyArray<ReadonlyArray<TableCellData>>;
/**
 * Reads the table-style GUID from `<a:tbl><a:tblPr><a:tableStyleId>`.
 * The reference desktop app uses GUIDs to reference built-in table styles
 * (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` = "Medium Style 2 -
 * Accent 1", etc.) and theme-local styles. Returns `null` when the
 * table doesn't reference one (uses the slide's default style).
 */
declare const getTableStyleId: (table: SlideShapeData) => string | null;
/**
 * Applies a table style by writing `<a:tbl><a:tblPr><a:tableStyleId>`.
 * `styleId` is a style GUID with its curly braces — e.g.
 * `'{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'` — or the English name of one
 * of the reference desktop app's built-in styles (`'Medium Style 2 - Accent 1'`; see
 * `BUILTIN_TABLE_STYLES`). For a built-in style, the reference desktop app's definition is
 * added to `ppt/tableStyles.xml` (creating the part if needed) unless the
 * deck already defines that GUID, as the reference desktop app does when you pick a style.
 * Pass `null` to remove the element so the table falls back to the deck's
 * default style. Creates `<a:tblPr>` if absent. Throws when the shape isn't
 * a table graphic frame or `styleId` is neither a GUID nor a built-in name.
 */
declare const setTableStyleId: (table: SlideShapeData, styleId: BuiltinTableStyleName | (string & {}) | null) => void;
/**
 * Reads the boolean style flags off `<a:tblPr>` — which header /
 * footer rows + columns are banded or emphasised. Mirrors the
 * `firstRow` / `bandRow` flags exposed by `addSlideTable`.
 *
 * Per ECMA-376 §17.18.95 / §21.1.3.15:
 *
 *   - `firstRow` — header row is styled differently.
 *   - `lastRow` — total row is styled differently.
 *   - `firstCol` — first column is styled differently.
 *   - `lastCol` — last column is styled differently.
 *   - `bandRow` — alternating row shading.
 *   - `bandCol` — alternating column shading.
 *
 * Renderers use these to switch on the corresponding table style
 * variant. Returns all-`false` for tables that don't author `<a:tblPr>`.
 */
declare const getTableStyleFlags: (table: SlideShapeData) => {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
};
/**
 * Sets one or more boolean style flags on `<a:tblPr>`. Only the keys
 * present in `flags` are touched — omitted keys are left at their
 * current state. A flag set to `false` strips the attribute (rather
 * than emitting `="0"`), matching how the reference desktop app round-trips defaults.
 *
 * See `getTableStyleFlags` for the meaning of each flag.
 */
declare const setTableStyleFlags: (table: SlideShapeData, flags: {
  firstRow?: boolean;
  lastRow?: boolean;
  firstCol?: boolean;
  lastCol?: boolean;
  bandRow?: boolean;
  bandCol?: boolean;
}) => void;
/**
 * Returns the table's row + column counts. Throws when the shape
 * isn't a table graphic frame.
 */
declare const getTableDimensions: (table: SlideShapeData) => {
  readonly rows: number;
  readonly cols: number;
};
/**
 * Returns each column's width in EMU, in left-to-right order, as
 * declared on `<a:tblGrid>/<a:gridCol w="...">`. Missing or
 * unparseable widths default to 0. Throws when the shape isn't a
 * table graphic frame.
 */
declare const getTableColumnWidths: (table: SlideShapeData) => ReadonlyArray<Emu>;
/**
 * Returns the table's nominal `(width, height)` derived from
 * summing the `<a:gridCol w>` and `<a:tr h>` attributes (both in
 * EMU). Useful for layout pipelines that want to know how big a
 * table really is without dereferencing the shape's `<a:xfrm>`.
 *
 * Throws when the shape isn't a table graphic frame.
 */
declare const getTableSize: (table: SlideShapeData) => {
  readonly width: Emu;
  readonly height: Emu;
};
/**
 * Returns each row's height in EMU, in top-to-bottom order, from
 * `<a:tr h="...">`. Missing or unparseable heights default to 0.
 * Throws when the shape isn't a table graphic frame.
 */
declare const getTableRowHeights: (table: SlideShapeData) => ReadonlyArray<Emu>;
/**
 * Sets a single column's width on the table grid. Throws on
 * out-of-range column indices or non-table shapes. The total table
 * width is not auto-adjusted — callers are responsible for keeping
 * the sum consistent with the table's `<a:xfrm>` extent if the reference desktop app
 * is to render the table without clipping.
 */
declare const setTableColumnWidth: (table: SlideShapeData, col: number, width: Emu) => void;
/**
 * Sets a single row's height. Throws on out-of-range row indices or
 * non-table shapes. As with `setTableColumnWidth`, the table's
 * `<a:xfrm>` extent is left to the caller.
 */
declare const setTableRowHeight: (table: SlideShapeData, row: number, height: Emu) => void;
/**
 * Returns the cell at `(row, col)`. Throws on out-of-range coordinates
 * or non-table shapes.
 */
declare const getTableCell: (table: SlideShapeData, row: number, col: number) => TableCellData;
/**
 * Replaces a cell's text. `\n` starts a new paragraph unless `newlines: 'break'`
 * requests paragraph-internal `<a:br>` elements. The paragraph-end
 * format (`<a:endParaRPr>`) is not kept unless `preserveFormatting` is enabled.
 * That option preserves unaffected runs and paragraph properties during editing.
 * With `range`, `text` replaces exactly that UTF-16 selection, always preserving
 * the formatting of unaffected text. Pass `{ case: 'upper' | 'lower' | 'sentence' |
 * 'title' | 'toggle' }` to change case in the whole cell or `options.range`
 * without rebuilding its runs, fields, or paragraph properties.
 */
declare const setTableCellText: (cell: TableCellData, text: string | {
  case: TextCase;
}, options?: {
  newlines?: "paragraph" | "break";
  preserveFormatting?: boolean;
  range?: {
    start: number;
    end: number;
  };
}) => void;
/**
 * Replaces the cell's text with explicitly structured paragraphs and runs —
 * the table counterpart of `setShapeParagraphs`. Read back with
 * `getTableCellParagraphs`.
 */
declare const setTableCellParagraphs: (cell: TableCellData, paragraphs: ReadonlyArray<ParagraphSpec>) => void;
/**
 * Reads the cell's merge / span attributes:
 *
 *   - `gridSpan` — number of columns this cell spans (≥1, default 1).
 *   - `rowSpan` — number of rows this cell spans (≥1, default 1).
 *   - `hMerge` — `true` when this cell is the right half of a horizontal
 *     span (it's painted by an earlier cell with `gridSpan > 1`).
 *   - `vMerge` — `true` when this cell is the bottom half of a vertical
 *     span (it's painted by an earlier cell with `rowSpan > 1`).
 *
 * Renderers should skip painting cells where `hMerge` or `vMerge` is
 * true; those cells exist only so the row/column grid stays consistent.
 */
declare const getTableCellSpan: (cell: TableCellData) => {
  gridSpan: number;
  rowSpan: number;
  hMerge: boolean;
  vMerge: boolean;
};
/**
 * Merges a rectangular block of table cells into a single visual cell.
 *
 * The block's top-left cell `(row, col)` becomes the anchor and carries
 * `gridSpan` (= `colSpan`) / `rowSpan`; the cells it covers are marked
 * `hMerge` / `vMerge` per ECMA-376 §21.1.3.18 (`CT_TableCell`) so the
 * grid stays rectangular while the reference desktop app paints only the anchor. This
 * is the write counterpart to {@link getTableCellSpan}.
 *
 * Constraints, enforced loudly (these are authoring-boundary inputs):
 *
 *   - `rowSpan` / `colSpan` must be ≥ 1, and at least one must be > 1
 *     (a 1×1 "merge" is a no-op the caller didn't mean).
 *   - The block must lie fully inside the table grid.
 *   - No cell in the block may already participate in another merge —
 *     overlapping merges corrupt the grid and trip the reference desktop app's repair
 *     dialog. Split the existing merge first.
 *
 * The anchor cell's text is preserved. `coveredText: 'keep'` (the default)
 * leaves each covered cell's `<a:txBody>` in the XML; `'drop'` removes it,
 * so the covered cells carry no `<a:txBody>` at all (CT_TableCell allows
 * that), which is how PptxGenJS writes a merge; `getTableCellParagraphs`
 * reads them as `[]`. `'append'` moves covered paragraphs into the anchor in
 * row-major order, preserving their formatting and leaving covered cells empty.
 */
declare const mergeTableCells: (table: SlideShapeData, block: {
  readonly row: number;
  readonly col: number;
  readonly rowSpan: number;
  readonly colSpan: number;
}, options?: {
  readonly coveredText?: "keep" | "drop" | "append";
}) => void;
/**
 * Splits the merged region containing `cell`, including a covered cell.
 * Retains each cell's text and properties; text removed during merging stays empty.
 * An unmerged cell is unchanged. Malformed regions throw before any mutation.
 */
declare const splitTableCell: (cell: TableCellData) => void;
/**
 * One side of a cell's border, as read from `<a:tcPr><a:ln{L|R|T|B}>`.
 * `widthEmu` is the line width in EMU; `color` is `#RRGGBB` or `null`
 * when the border is inherited / un-styled.
 */
interface TableCellBorder {
  readonly color: string | null;
  readonly widthEmu: number | null;
  readonly dash: string | null;
}
interface TableCellBorders {
  readonly left: TableCellBorder | null;
  readonly right: TableCellBorder | null;
  readonly top: TableCellBorder | null;
  readonly bottom: TableCellBorder | null;
  readonly tlToBr: TableCellBorder | null;
  readonly blToTr: TableCellBorder | null;
}
/**
 * Reads the per-side borders on a cell. Returns `null` for sides with
 * no explicit `<a:ln{Side}>` element (those inherit from the table
 * style / theme). All four cardinal sides plus the two diagonals
 * (TL→BR, BL→TR) are surfaced because real templates use them all.
 */
declare const getTableCellBorders: (pres: PresentationData, cell: TableCellData) => TableCellBorders;
/**
 * Sets one or more side borders on a cell. Sides listed with `null` are
 * removed from the `<a:tcPr>`; sides omitted from the partial are left
 * untouched. Pass the whole `sides` arg as `null` to clear every side
 * in one call. Creates `<a:tcPr>` if absent.
 *
 * The diagonal sides (`tlToBr` / `blToTr`) are independent of the
 * cardinal four — a strikethrough cell can have only a diagonal.
 */
declare const setTableCellBorders: (cell: TableCellData, sides: {
  left?: Partial<TableCellBorder> | null;
  right?: Partial<TableCellBorder> | null;
  top?: Partial<TableCellBorder> | null;
  bottom?: Partial<TableCellBorder> | null;
  tlToBr?: Partial<TableCellBorder> | null;
  blToTr?: Partial<TableCellBorder> | null;
} | null) => void;
/**
 * Reads the cell's text direction (`<a:tcPr vert="…"/>`) — same tokens
 * as `getShapeTextDirection`. Returns `null` for the default
 * horizontal direction.
 *
 * Vertical column headers in tables almost always emit `<a:tcPr
 * vert="vert270"/>` or `"eaVert"` so the header label reads bottom-to-
 * top alongside its column.
 */
declare const getTableCellTextDirection: (cell: TableCellData) => "vert" | "vert270" | "wordArtVert" | "eaVert" | "mongolianVert" | "wordArtVertRtl" | null;
/**
 * Reads the 3-D on the cell's text body (`<a:scene3d>` / `<a:sp3d>` under its
 * `<a:bodyPr>`) — the table-cell counterpart of `getShapeText3D`. Returns
 * `null` when the body has neither.
 */
declare const getTableCellText3D: (cell: TableCellData) => ReadText3D | null;
/**
 * Sets the cell's text-direction via `<a:tcPr vert="…"/>`. Same six
 * `ST_TextVerticalType` values as `setShapeTextDirection`. Passing
 * `null` (or `'horz'`) clears the attribute so the cell uses the
 * default horizontal direction. Creates `<a:tcPr>` if absent.
 */
declare const setTableCellTextDirection: (cell: TableCellData, direction: "horz" | "vert" | "vert270" | "wordArtVert" | "eaVert" | "mongolianVert" | "wordArtVertRtl" | null) => void;
/**
 * Reads the cell's vertical text anchor (`<a:tcPr anchor="t|ctr|b"/>`)
 * — `'top'`, `'center'`, `'bottom'`, or `null` for the default
 * (`ctr` / center per the schema).
 */
declare const getTableCellAnchor: (cell: TableCellData) => "top" | "center" | "bottom" | null;
/**
 * Sets the cell's vertical text anchor (`<a:tcPr anchor="t|ctr|b"/>`).
 * Pass `null` to clear the attribute so the cell uses the table's
 * default (`ctr`). Creates `<a:tcPr>` if absent.
 */
declare const setTableCellAnchor: (cell: TableCellData, anchor: "top" | "center" | "bottom" | null) => void;
/**
 * Reads the cell's inset margins (`<a:tcPr marL marR marT marB>`) in
 * EMU. Each side is `null` when the cell doesn't author it (renderers
 * should fall back to the reference desktop app's defaults — 91440 EMU / 0.1 inch
 * for the horizontal margins, 45720 EMU for the vertical).
 */
declare const getTableCellMargins: (cell: TableCellData) => {
  left: number | null;
  right: number | null;
  top: number | null;
  bottom: number | null;
};
/**
 * Sets the cell's inset margins (`<a:tcPr marL marR marT marB>`) in
 * EMU. Sides set to `null` are stripped from the XML so they fall back
 * to the reference desktop app's defaults. Creates `<a:tcPr>` if absent.
 *
 * Pass `null` for the whole `margins` arg to drop every authored side
 * at once.
 */
declare const setTableCellMargins: (cell: TableCellData, margins: {
  left?: number | null;
  right?: number | null;
  top?: number | null;
  bottom?: number | null;
} | null) => void;
/** Reads the cell's plain text (paragraphs joined with `\n`). */
declare const getTableCellText: (cell: TableCellData) => string;
/** One paragraph of a cell's text, with its alignment and inline elements. */
interface TableCellParagraph {
  /**
   * Horizontal alignment from `<a:pPr algn>`, or `null` when unset (the
   * cell then inherits the reference desktop app's left default).
   */
  readonly align: ParagraphAlignment | null;
  /** Runs / fields / breaks in document order, with their literal `<a:rPr>` format. */
  readonly elements: ReadonlyArray<ShapeParagraphElement & {
    readonly clickAction?: ShapeClickAction;
    readonly tooltip?: string;
  }>;
  /**
   * Literal format of the paragraph-end mark (`<a:endParaRPr>`), or `null`
   * when absent — the only format a paragraph with no `elements` carries.
   */
  readonly endFormat: ReadTextFormat | null;
}
/**
 * Reads a cell's text as structured paragraphs — each carrying its
 * alignment and per-run format (`size`, `bold`, `italic`, `color`,
 * `font`, …) — for renderers that need to reproduce styled cell text.
 * `getTableCellText` is the flat-string counterpart when only the visible
 * characters matter. Returns an empty array when the cell has no
 * `<a:txBody>`.
 *
 * The cell's `<a:txBody>` uses the same DrawingML paragraph grammar as a
 * shape's, so this shares `readParagraphElements` with
 * `getShapeParagraphElements`. Table-style and theme-level defaults are not
 * folded in here: explicit run properties win, and absent ones are left
 * `undefined` for the caller to resolve.
 */
declare const getTableCellParagraphs: (cell: TableCellData) => ReadonlyArray<TableCellParagraph>;
/**
 * Reads a table-cell run's effective character format. The literal
 * `getTableCellParagraphs` format only contains properties authored on the
 * run; this resolver also applies the cell text body's paragraph defaults and
 * outline-level defaults (`a:pPr/defRPr` and `a:lstStyle`). Pass `null` as
 * `runIndex` to resolve the paragraph end mark, including empty paragraphs.
 * Pass `{ fieldIndex }` for a zero-based `a:fld` index; numeric indices count
 * regular `a:r` runs only. Pass `{ breakIndex }` for a zero-based `a:br` index.
 */
declare const getTableCellRunFormatEffective: (pres: PresentationData, cell: TableCellData, paragraphIndex: number, runIndex: number | null | {
  readonly fieldIndex: number;
} | {
  readonly breakIndex: number;
}) => ReadTextFormat;
/** Sets a solid background color on a cell (`<a:tcPr><a:solidFill>`). */
declare const setTableCellFill: (cell: TableCellData, color: Color) => void;
/** Removes any background fill from a cell. */
declare const clearTableCellFill: (cell: TableCellData) => void;
/**
 * Reads the cell's solid background color. Returns `#RRGGBB` for
 * `<a:srgbClr>`, `scheme:<token>` for `<a:schemeClr>`, or `null` when
 * the cell has no fill, no `<a:tcPr>`, or the fill is non-solid
 * (gradient / pattern / image).
 */
declare const getTableCellFill: (cell: TableCellData) => string | null;
/** Effective cell appearance after table-style and cell-property inheritance. */
interface TableCellAppearanceEffective {
  readonly fill: ShapeFill;
  /** Opacity (0–1) of a solid `fill` whose color carries an alpha transform. */
  readonly fillOpacity?: number;
  readonly borders: TableCellBorders;
}
/** Resolves table-style `tcStyle` fills and borders, then overlays local `tcPr`. */
declare const getTableCellAppearanceEffective: (pres: PresentationData, cell: TableCellData) => TableCellAppearanceEffective;
/** Effective background drawn behind a table's cells. */
interface TableBackgroundEffective {
  readonly fill: ShapeFill;
  /** Opacity (0–1) of a solid `fill` whose color carries an alpha transform. */
  readonly fillOpacity?: number;
  /** Stops (with theme-resolved colors) and direction of a gradient `fill`. */
  readonly gradient?: ReadGradientFill;
}
/**
 * Resolves the table background: a fill on `<a:tblPr>` itself, else the
 * table style's `<a:tblBg>` (the reference desktop app's Themed Styles use a theme
 * style-matrix fill there). Cells without a fill of their own show it.
 * Returns `null` when neither defines one. The `tblBg` effect is not read.
 */
declare const getTableBackgroundEffective: (pres: PresentationData, table: SlideShapeData) => TableBackgroundEffective | null;
/** Applies a TextFormat to the cell's text, optionally within UTF-16 offsets
 * in getTableCellText (exclusive end). Breaks count as one character; invalid
 * ranges and split-surrogate boundaries throw without changing the text.
 * `reset` restores inherited run appearance before applying `format`, retaining
 * links and language. Without either target, run-format defaults are also cleared.
 * Pass a zero-based `paragraphEnd` instead of `range` to update a paragraph's end mark. */
declare const setTableCellTextFormat: (cell: TableCellData, format: TextFormat, options?: {
  range?: {
    start: number;
    end: number;
  };
  reset?: boolean;
  paragraphEnd?: never;
} | {
  paragraphEnd: number;
  reset?: boolean;
  range?: never;
}) => void;
/**
 * Sets or removes a click link on cell text. Omitting range covers the whole
 * cell; otherwise offsets are UTF-16 positions in getTableCellText, with an
 * exclusive end. Formatting and links outside the range are preserved.
 * Omitting tooltip clears the selected link's previous description.
 */
declare const setTableCellClickAction: (cell: TableCellData, action: ShapeClickAction | null, options?: {
  range?: {
    start: number;
    end: number;
  };
  tooltip?: string;
}) => void;
/** Sets horizontal alignment on every paragraph in the cell. */
declare const setTableCellAlignment: (cell: TableCellData, align: ParagraphAlignment) => void;
/** Zero-based (row, col) of the cell. */
declare const getTableCellPosition: (cell: TableCellData) => {
  row: number;
  col: number;
};
/**
 * Reads the horizontal alignment from the cell's first paragraph as its
 * spec token (`l`, `ctr`, `r`, `just`, `dist`, `justLow`, `thaiDist`) — so
 * `setTableCellAlignment(cell, 'center')` reads back as `'ctr'`. Returns
 * `null` when the cell has no `<a:txBody>`, or its first paragraph has no
 * valid `algn` attribute (the reference desktop app then defaults to `l`). For a
 * plain-English name per paragraph, use `getTableCellParagraphs`.
 */
declare const getTableCellAlignment: (cell: TableCellData) => ParagraphAlignmentToken | null;
/**
 * Inserts a row into the table. `atIndex` is 0-based; `undefined`
 * appends at the end. `cells` supplies cell values; missing cells
 * become blank, extras are dropped. The row's height matches the
 * average of existing rows (or a 0.4in default for empty tables).
 */
declare const insertTableRow: (table: SlideShapeData, atIndex?: number, cells?: ReadonlyArray<string>) => void;
/** Removes the row at `atIndex` from the table. Throws on out-of-range. */
declare const removeTableRow: (table: SlideShapeData, atIndex: number) => void;
/**
 * Inserts a column into the table. `atIndex` defaults to the end.
 * `widthEmu` defaults to the average of existing column widths (or
 * 914400 = 1in if the table has no columns). Existing rows get a new
 * blank cell at `atIndex`.
 */
declare const insertTableColumn: (table: SlideShapeData, atIndex?: number, widthEmu?: number) => void;
/** Removes the column at `atIndex` (and the corresponding cell in every row). */
declare const removeTableColumn: (table: SlideShapeData, atIndex: number) => void;
//#endregion
//#region src/internal/validator/validate-presentation.d.ts
type IssueSeverity = 'error' | 'warning';
interface ValidationIssue {
  readonly severity: IssueSeverity;
  readonly message: string;
  readonly partName?: PartName;
}
//#endregion
//#region src/api/fn/validation.d.ts
/**
 * Runs a set of lightweight invariant checks on the package and
 * returns the list of issues found. An empty array means the deck
 * passes every check.
 *
 * Catches the common authoring mistakes — missing presentation.xml,
 * dangling slide rels, slides without a layout, etc. — without
 * depending on a heavyweight XSD engine, so it runs identically in
 * Node and the browser.
 *
 * Use it as a pre-save sanity check, especially after orchestrating
 * lots of mutations against the same package. Higher-fidelity XSD
 * validation lives in the test harness (Layer 1) and stays Node-only.
 */
declare const validatePresentation: (pres: PresentationData) => ReadonlyArray<ValidationIssue>;
//#endregion
//#region src/api/fn/package-introspection.d.ts
/**
 * Power-user escape hatch. Returns the underlying `OpcPackage`
 * backing `pres`. Use this when you need to manipulate parts /
 * rels directly. Most callers should use the typed helpers above
 * (`listPackageParts`, `readPackagePart`, `getMediaParts`, etc.).
 *
 * @internal — used by `@office-kit/pptx/node` to mount fs-backed helpers.
 */
declare const _internalPackageOf: (pres: PresentationData) => OpcPackage;
/** One entry in the package's parts list. */
interface PackagePartInfo {
  readonly name: string;
  readonly contentType: string;
  readonly byteLength: number;
}
/**
 * Enumerates every OPC part in the package. Useful for advanced
 * inspection (e.g. "what parts does this template carry?") without
 * dropping to `_internalPackageOf`.
 */
declare const listPackageParts: (pres: PresentationData) => ReadonlyArray<PackagePartInfo>;
/**
 * Reads a single OPC part's bytes by part name (e.g.
 * `'/ppt/slides/slide1.xml'`). Returns `null` when no such part
 * exists. The returned `Uint8Array` is a live view into the
 * package — DO NOT mutate it. Use this for read-only inspection
 * (e.g. parsing custom extension parts).
 */
declare const readPackagePart: (pres: PresentationData, name: string) => Uint8Array | null;
/** A media (image / video / audio) part embedded in the package. */
interface MediaPart {
  readonly name: string;
  readonly contentType: string;
  readonly data: Uint8Array;
}
/**
 * Returns the total size of the package's parts in bytes
 * (uncompressed). Useful for storage estimation, quota checks,
 * and "how big is this deck before save?" diagnostics. The
 * actual `savePresentation` output is typically smaller after
 * DEFLATE; this is an upper bound.
 */
declare const getPackageSize: (pres: PresentationData) => number;
/**
 * Returns every `/ppt/media/...` part in the package. Useful for
 * audit / export workflows — e.g. "extract every embedded image."
 */
declare const getMediaParts: (pres: PresentationData) => ReadonlyArray<MediaPart>;
/**
 * Returns every media part NOT referenced by any rels in the
 * package — the set `compactPackage` would remove. Non-destructive;
 * the caller decides whether to delete.
 *
 * Useful for audit UIs that want to surface bloat before cleaning,
 * and for "is this asset still used?" checks.
 */
declare const getOrphanMediaPartNames: (pres: PresentationData) => ReadonlyArray<string>;
/**
 * Returns every media part name the slide's rels reference
 * (typically `/ppt/media/imageN.ext`). Walks the slide's rels
 * graph and resolves each internal target. Useful for "which
 * media files does this slide depend on?" audits.
 */
declare const getSlideMediaPartNames: (slide: SlideData) => ReadonlyArray<string>;
/**
 * Returns every slide that references the given media part name
 * (typically `/ppt/media/imageN.ext`). Walks each slide's rels and
 * checks whether any internal rel resolves to `mediaPartName`.
 *
 * Useful for image-audit workflows: "before I replace this image,
 * which slides will the change affect?"
 */
declare const slidesUsingMediaPart: (pres: PresentationData, mediaPartName: string) => ReadonlyArray<SlideData>;
/**
 * Removes media parts that no rels graph references. Returns the
 * list of removed part names. Useful after a sequence of slide
 * removals leaves orphan images behind.
 *
 * Only `/ppt/media/...` parts are considered. The check walks every
 * `.rels` part in the package and resolves each internal rel target
 * against its source part name to build the live media set.
 */
declare const compactPackage: (pres: PresentationData) => {
  readonly removed: ReadonlyArray<string>;
};
/**
 * Replaces the bytes of a media part in place. Returns `true` if the
 * part was found and updated, `false` otherwise. The content type is
 * preserved.
 *
 * Useful for the "swap every instance of this logo" workflow — pick
 * the right `partName` via `getMediaParts` and call this once. Every
 * `<a:blip r:embed="…"/>` reference is unaffected because the rels
 * already point at this part name.
 */
declare const setMediaPartBytes: (pres: PresentationData, partName: string, bytes: Uint8Array) => boolean;
/**
 * High-level snapshot of the presentation's structure. Useful as a
 * diagnostic checklist when debugging a template or generating audit
 * reports. The numbers reflect what's reachable through the typed
 * API on the current package state.
 */
interface PresentationSummary {
  readonly slideCount: number;
  readonly hiddenSlideCount: number;
  readonly totalShapes: number;
  readonly shapesByKind: Readonly<Record<ShapeKind, number>>;
  readonly layoutCount: number;
  readonly sectionCount: number;
  readonly partCount: number;
  readonly hasCharts: boolean;
  readonly hasComments: boolean;
  readonly hasAnimations: boolean;
  readonly themeName: string | null;
}
declare const getPresentationSummary: (pres: PresentationData) => PresentationSummary;
//#endregion
//#region src/api/fn/shape-image-effects.d.ts
/** A color used by a two-color image recolor, with optional DrawingML transforms. */
type ImageRecolorColor = Color | {
  readonly color: Color;
  readonly colorTransforms?: readonly ColorTransform[];
};
/**
 * An image recolor operation from the reference desktop app. Threshold is expressed as a percent
 * (0..100), matching `getShapeImageBiLevelThreshold` and the reference desktop app's UI.
 */
type ImageRecolor = {
  readonly kind: 'none';
} | {
  readonly kind: 'grayscale';
} | {
  readonly kind: 'duotone';
  readonly colors: readonly [ImageRecolorColor, ImageRecolorColor];
} | {
  readonly kind: 'threshold';
  readonly threshold: number;
} | {
  readonly kind: 'washout';
};
/**
 * Returns the embedded image bytes for a picture shape, or `null`
 * when the shape isn't a picture or has no `r:embed` reference
 * (external images aren't followed).
 *
 * The returned `Uint8Array` is a live view into the package media
 * part — treat it as read-only; copy if you need an independent
 * buffer.
 */
declare const getShapeImageBytes: (shape: SlideShapeData) => Uint8Array | null;
/**
 * `true` when the shape's text body carries any visible characters.
 * Tighter than checking `getShapeText(shape) !== ''` because it
 * doesn't allocate the concatenated string.
 */
declare const hasShapeText: (shape: SlideShapeData) => boolean;
/**
 * `true` when the shape carries an embedded image — either a
 * `<p:pic>` picture or a `<p:spPr>/<a:blipFill>` image fill on a
 * regular shape. External `r:link` references count too.
 */
declare const hasShapeImage: (shape: SlideShapeData) => boolean;
/**
 * Dense per-slide image count array. Counts every shape that
 * `hasShapeImage` matches on each slide. Rounds out the density-array
 * family alongside the chart, table, shape, text, and comment
 * counters.
 */
declare const getPresentationImageCountsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * Returns every shape on the slide that is mirrored — horizontally
 * (`flipH`), vertically (`flipV`), or both.
 */
declare const findFlippedShapes: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/**
 * Returns every unordered pair of shapes on the slide whose
 * bounding boxes overlap. Built on `shapesOverlap`. Pairs are
 * returned with `a` strictly preceding `b` in document order, and
 * each pair appears at most once.
 *
 * Useful for layout audits — "do any boxes collide on this slide?"
 * Shapes without `<a:xfrm>` bounds never overlap anything.
 */
declare const findOverlappingShapePairs: (slide: SlideData) => ReadonlyArray<readonly [SlideShapeData, SlideShapeData]>;
/**
 * Returns every shape on the slide whose bounding box extends past
 * the slide canvas (`getSlideSize(pres)`). Useful audit helper for
 * catching shapes that the reference desktop app will silently render off-screen
 * or clip on export. Shapes without `<a:xfrm>` bounds are skipped.
 *
 * If the presentation has no slide-size declared, every positioned
 * shape is returned (caller can't audit against an absent canvas).
 */
declare const findShapesOutsideCanvas: (slide: SlideData, pres: PresentationData) => ReadonlyArray<SlideShapeData>;
/**
 * Every slide whose layout's `cSld@name` matches the given string.
 * Useful for batch operations on slides sharing a layout — for
 * example, restyling every "Title and Content" slide in a deck.
 *
 * Matching is exact (case-sensitive). Slides without a resolved
 * layout are skipped.
 */
declare const findSlidesByLayoutName: (pres: PresentationData, layoutName: string | RegExp) => ReadonlyArray<SlideData>;
/**
 * Every slide whose resolved layout part name equals `layoutPartName`
 * (e.g. `'/ppt/slideLayouts/slideLayout3.xml'`). Stable across
 * template-name collisions and locale renames — keyed on the actual
 * package path. Pair to `findSlidesByLayoutName` /
 * `findSlidesByLayoutType` for cases where the caller already has a
 * layout part name in hand (typical when iterating over a layouts
 * collection programmatically).
 */
declare const findSlidesByLayoutPartName: (pres: PresentationData, layoutPartName: string) => ReadonlyArray<SlideData>;
/**
 * Every slide whose layout `@type` (e.g. `'title'`, `'blank'`,
 * `'obj'`) matches. Sibling of `findSlidesByLayoutName`, but keyed
 * on the OOXML layout-type enum rather than the human-facing name —
 * stable across locales and template providers.
 */
declare const findSlidesByLayoutType: (pres: PresentationData, layoutType: SlideLayoutType | string) => ReadonlyArray<SlideData>;
/**
 * Returns the package part name (`/ppt/media/imageN.ext`) of
 * whichever image the shape carries — picture (`<p:pic>`) or
 * image-fill (`<a:blipFill>` nested under `<p:spPr>`). Returns
 * `null` when the shape has no embedded image, or the rel points
 * at an external `r:link` rather than an internal target.
 *
 * Useful for addressing the media part directly with
 * `setMediaPartBytes` or `readPackagePart`.
 */
/**
 * Returns the external URL of the picture when its `<a:blip>` carries an
 * `r:link` (external) relationship rather than an `r:embed`. Returns
 * `null` for embedded pictures, non-picture shapes, or when the
 * relationship doesn't resolve.
 *
 * The reference desktop app emits `r:link` when the user inserts via "Link to file"
 * instead of "Insert Picture". The bytes live outside the package, so
 * `getShapeImageBytes` can't render them — readers / preview tools
 * should fall back to this URL or a placeholder.
 */
/**
 * Returns `true` when the picture's `<a:blip>` carries `<a:grayscl/>`
 * — the reference desktop app's "Color > Grayscale" recolor preset. Renderers can
 * project this onto a CSS `filter: grayscale(100%)` or an SVG
 * `<feColorMatrix>` desaturation.
 */
declare const isShapeImageGrayscale: (shape: SlideShapeData) => boolean;
/**
 * Returns the threshold of the picture's `<a:blip><a:biLevel thresh="…"/>`
 * effect — the reference desktop app's "Color > Black and White" preset. Threshold is
 * a percent (0..100); pixels brighter become white, darker become black.
 * Returns `null` when no biLevel transform is set.
 */
declare const getShapeImageBiLevelThreshold: (shape: SlideShapeData) => number | null;
/**
 * Reads the picture's duotone color transform from `<a:blip><a:duotone>`.
 * The reference desktop app emits two `<a:srgbClr>` (or scheme color) children for a
 * two-color duotone effect — typical "Picture Tools › Recolor".
 * By default colors are resolved to RGB for rendering. Pass
 * `{ resolveColors: false }` to preserve authoring color models and
 * transforms for a read-edit-write cycle. Only sRGB and scheme colors can
 * retain their authoring model; other DrawingML models are returned as their
 * resolved base color with the original transforms. Returns `null` when no
 * duotone is set.
 */
declare function getShapeImageDuotone(pres: PresentationData, shape: SlideShapeData): {
  firstColor: string | null;
  secondColor: string | null;
} | null;
declare function getShapeImageDuotone(pres: PresentationData, shape: SlideShapeData, options: {
  readonly resolveColors: false;
}): Extract<ImageRecolor, {
  kind: 'duotone';
}> | null;
declare function getShapeImageDuotone(pres: PresentationData, shape: SlideShapeData, options: {
  readonly resolveColors: true;
}): {
  firstColor: string | null;
  secondColor: string | null;
} | null;
declare function getShapeImageDuotone(pres: PresentationData, shape: SlideShapeData, options: {
  readonly resolveColors: boolean;
}): {
  firstColor: string | null;
  secondColor: string | null;
} | Extract<ImageRecolor, {
  kind: 'duotone';
}> | null;
declare const getShapeImageLinkUrl: (shape: SlideShapeData) => string | null;
declare const getShapeImagePartName: (shape: SlideShapeData) => string | null;
/**
 * Returns the bytes of the image used as this shape's *fill*
 * (`<a:blipFill>` nested under `<p:spPr>`, as written by
 * `setShapeImageFill`). Distinct from `getShapeImageBytes`, which only
 * applies to `<p:pic>` picture shapes.
 *
 * Returns null if the shape has no image fill, the blip has no
 * `r:embed`, or the embed points at an external `r:link`.
 */
declare const getShapeImageFillBytes: (shape: SlideShapeData) => Uint8Array | null;
/**
 * Returns the image format token (`'png'`, `'jpeg'`, …) for whichever
 * image bytes the shape carries — picture (`<p:pic>`) or image-fill
 * (`<a:blipFill>` on `<p:spPr>`). Returns `null` if the shape has no
 * embedded image or the bytes don't match a recognized signature.
 */
declare const getShapeImageFormat: (shape: SlideShapeData) => ImageFormat | null;
/** One of the reference desktop app's Artistic Effects ([MS-ODRAWXML] CT_PictureEffect, `a14:artistic*`). */
type ImageArtisticEffect = 'marker' | 'pencilGrayscale' | 'pencilSketch' | 'lineDrawing' | 'chalkSketch' | 'paintStrokes' | 'paintBrush' | 'glowDiffused' | 'blur' | 'lightScreen' | 'watercolorSponge' | 'filmGrain' | 'mosaicBubbles' | 'glass' | 'cement' | 'texturizer' | 'crisscrossEtching' | 'pastelsSmooth' | 'plasticWrap' | 'cutout' | 'photocopy' | 'glowEdges';
/**
 * Reads the Artistic Effect the reference desktop app applied to a picture or image fill, or
 * `null` when there is none.
 *
 * The reference desktop app stores the effect in the blip's `a14:imgProps` extension next to
 * a relationship to the original picture (JPEG XR). The embedded picture is
 * already the effect's result, so renderers draw it as is.
 */
declare const getShapeImageArtisticEffect: (shape: SlideShapeData) => ImageArtisticEffect | null;
/**
 * Reads the picture or image fill's opacity (0–1 fraction). Returns `null` when no
 * `<a:alphaModFix>` is present (the reference desktop app treats absence as fully
 * opaque); returns `1` when an explicit alphaModFix sets full opacity.
 */
declare const getShapeImageOpacity: (shape: SlideShapeData) => number | null;
/**
 * Reads a picture or image fill's crop fractions. Returns `null` when no
 * `<a:srcRect>` is present; otherwise returns a fully-populated object
 * with every side filled in (0 for omitted sides on disk).
 */
declare const getShapeImageCrop: (shape: SlideShapeData) => ImageCrop | null;
/**
 * Adjusts a picture or image fill's brightness via `<a:blip><a:lum bright="…"/>`. The value
 * is a -1..1 fraction:
 *
 *   - `1`     → +100% brightness
 *   - `0` or `null` → no brightness change (the `bright` attribute is removed)
 *   - `-1`    → -100% brightness
 *
 * Brightness and contrast share the one `<a:lum>` element, so setting one keeps
 * the other. Throws for shapes without an image and on values outside [-1, 1].
 */
declare const setShapeImageBrightness: (shape: SlideShapeData, value: number | null) => void;
/**
 * Adjusts a picture or image fill's contrast via `<a:blip><a:lum contrast="…"/>`. The value
 * is a -1..1 fraction (ECMA-376 `ST_FixedPercentage`):
 *
 *   - `0` or `null` → no contrast change (the `contrast` attribute is removed)
 *   - `0.5`         → +50% contrast (boosted)
 *   - `-0.5`        → -50% contrast (washed out)
 *
 * Brightness and contrast share the one `<a:lum>` element, so setting one keeps
 * the other. Throws for shapes without an image and on values outside [-1, 1].
 */
declare const setShapeImageContrast: (shape: SlideShapeData, value: number | null) => void;
/**
 * Reads a picture or image fill's contrast (the `<a:lum contrast>` fraction in [-1, 1]).
 * Returns `null` when no contrast adjustment is present.
 */
declare const getShapeImageContrast: (shape: SlideShapeData) => number | null;
/**
 * Reads a picture or image fill's brightness (the `<a:lum bright>` fraction in [-1, 1]).
 * Returns `null` when no brightness adjustment is present.
 */
declare const getShapeImageBrightness: (shape: SlideShapeData) => number | null;
/**
 * Clears the reference desktop app's image color corrections from a picture or image fill.
 * This removes grayscale (`grayscl`), duotone, bi-level recolor, and
 * brightness/contrast (`lum`) effects while preserving opacity, media
 * references, and all other known or extension effects.
 */
declare const resetShapeImageColorEffects: (shape: SlideShapeData) => void;
/**
 * Applies an image recolor from the reference desktop app to a picture or image-filled shape.
 * `none` removes the recolor while preserving opacity, media references, and
 * unrelated DrawingML effects. Threshold is a percentage from 0 to 100.
 */
declare const setShapeImageRecolor: (shape: SlideShapeData, recolor: ImageRecolor) => void;
/**
 * Sets picture or image-fill opacity (0–1; `1` is fully opaque).
 * Pass `null` to restore the reference desktop app's default opacity.
 * Rejects shapes without an image and values outside `[0, 1]` without changing them.
 */
declare const setShapeImageOpacity: (shape: SlideShapeData, opacity: number | null) => void;
/** Crop a picture by fraction of each side. Omitted sides default to 0. */
interface ImageCrop {
  readonly left?: number;
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
}
/**
 * Sets (or clears) a `<a:srcRect>` on a picture or image-filled shape, cropping the
 * embedded image by the given fraction on each side. Pass `null` to
 * remove an existing crop.
 *
 * Fractions use the signed `ST_Percentage` range per side. `{ left: 0.25 }`
 * clips 25% off the left edge; negative values extend the source rectangle.
 * Stretch fills fit the remaining image to the frame;
 * tiled fills repeat the remaining image. The shape's geometry (`<a:xfrm>`) is unchanged.
 */
declare const setShapeImageCrop: (shape: SlideShapeData, crop: ImageCrop | null) => void;
//#endregion
//#region src/api/fn/_animation-timing.d.ts
/** When a step runs relative to the one before it. */
type AnimationStart = 'click' | 'withPrevious' | 'afterPrevious' | 'unknown';
/**
 * What becomes of the value an effect animates once it has run, from the
 * `fill` on its behaviours (ST_TLTimeNodeFillType: `remove | freeze | hold |
 * transition`). `'unstated'` is not `'held'`: the schema gives the attribute
 * no default, so a tree that omits it has not said.
 */
type AnimationValueAfterEnd = 'held' | 'removed' | 'unstated';
/**
 * What an effect animates. `unsupported` covers targets this library does not
 * model — charts, OLE parts, sub-shapes, and composite effects whose
 * behaviours do not all animate the same thing. The step is still listed so a
 * caller can show it, but it cannot be edited.
 */
type AnimationTarget = {
  readonly kind: 'shape';
  readonly shapeId: number;
} | {
  readonly kind: 'paragraphs';
  readonly shapeId: number;
  readonly firstParagraph: number;
  readonly lastParagraph: number;
} | {
  readonly kind: 'unsupported';
  readonly shapeId: number | null;
};
/** One animation effect on a slide, as read from `<p:timing>`. */
interface SlideAnimationStep {
  /**
   * How the editing APIs address this step: the effect's `<p:cTn id>`.
   * `null` when the node carries no id, the id is not a plain integer, or the
   * same id appears more than once in the tree — none of which is a safe
   * handle, so such a step is read-only.
   */
  readonly id: number | null;
  readonly target: AnimationTarget;
  /**
   * Every shape any of the step's behaviours animates, in document order.
   * A composite effect that drives two shapes reports both here even though
   * `target` refuses to name one of them as *the* target.
   */
  readonly targetShapeIds: readonly number[];
  /** `null` for a preset outside the `AnimationEffect` tokens. */
  readonly effect: AnimationEffect | null;
  /**
   * Which edge (or corner) of the slide a fly, wipe or peek comes from or
   * leaves by, or which corner strips start from. `null` for every other
   * effect — the preset states no direction, so neither does this. The same
   * goes for the four options below, each `null` unless the effect takes it.
   */
  readonly direction: AnimationDirection | null;
  /** Blinds, checkerboard, random bars and split: which way the bars run. */
  readonly orientation: AnimationOrientation | null;
  /** Shape and split: whether the effect closes in on the centre or opens out of it. */
  readonly inOut: AnimationInOut | null;
  /** Shape: the outline the object is revealed or hidden through. */
  readonly shape: AnimationShape | null;
  /** Wheel: how many spokes. */
  readonly spokes: number | null;
  /** Spin: which way it turns. */
  readonly spinDirection: AnimationSpinDirection | null;
  /** Spin: how far it turns, in degrees. */
  readonly spinDegrees: number | null;
  /** Grow/Shrink: which axes it scales. */
  readonly scaleDirection: AnimationScaleDirection | null;
  /** Grow/Shrink: the size it reaches, in percent of the shape's own. */
  readonly scalePercent: number | null;
  /** Transparency: how transparent the shape becomes, 0–100. */
  readonly transparencyPercent: number | null;
  /**
   * The colour emphasis effects: the colour they change to — a theme slot as
   * `scheme:accent2`, an sRGB value as `#RRGGBB`, with its colour transforms
   * when it has any.
   */
  readonly color: AnimationColor | null;
  readonly presetId: number | null;
  readonly presetClass: string | null;
  /** `'unknown'` when the node type is not one this library models. */
  readonly start: AnimationStart;
  /** `null` when the tree does not state one — never defaulted. */
  readonly durationMs: number | null;
  /** `null` when the tree does not state one, or the delay is indefinite. */
  readonly delayMs: number | null;
  /**
   * What the tree says becomes of the value this effect animates once it has
   * run. It decides nothing on its own — `playable` already requires that the
   * behaviour saying whether the shape is on the slide keeps its value — but
   * it is what two effects animating one object at the same time turn on: the
   * one that ends first either stays above the other or gives way to it. A
   * player that finds `'unstated'` there must not choose for the file.
   */
  readonly valueAfterEnd: AnimationValueAfterEnd;
  /**
   * How the effect's build entry reveals the shape's text (`<p:bldP build>`):
   * as one object (`whole`, or no entry), all paragraphs at once
   * (`allAtOnce`), one paragraph at a time (`p`), or `'custom'` (`cust`) — a
   * build the reference desktop app's own Effect Options do not offer.
   */
  readonly build: AnimationTextBuild | 'custom';
  readonly buildLevel: number | null;
  /**
   * Which `<p:seq>` the step lives in, named after its `nodeType`. Only
   * `'mainSeq'` advances on the slide's own clicks: an `'interactiveSeq'`
   * sequence fires when the viewer clicks the shape it is bound to, so a
   * player that folded one into the click order would run it at the wrong
   * time — and swallow a click the slide owes its main sequence.
   */
  readonly sequence: AnimationSequenceKind;
  /**
   * `true` when a player can run this step as part of the slide's click
   * order: it is in the main sequence, starts in a way we model, animates a
   * single target we model, and uses a preset we know how to render. A step
   * that is not playable must be shown as such, never folded into the normal
   * progression.
   */
  readonly playable: boolean;
  /**
   * `true` when this library can safely move, remove or retime the step:
   * `playable`, plus a unique handle to address it by. Editability is not
   * playability — a step can be one without the other as either side grows.
   */
  readonly editable: boolean;
}
/**
 * Which `<p:seq>` an effect belongs to, named after its `nodeType`. The two
 * tokens are ST_TLTimeNodeType values; `'other'` covers a sequence whose node
 * type is absent or something else again.
 */
type AnimationSequenceKind = 'mainSeq' | 'interactiveSeq' | 'other';
//#endregion
//#region src/api/fn/shape-animation.d.ts
/**
 * Adds a single animation effect to the given shape.
 *
 * The effect is *merged* into any existing `<p:timing>` on the slide rather
 * than replacing it: animating a second shape (or re-running on a template that
 * already has authored animations) preserves the existing effects, with cTn ids
 * renumbered to stay unique. To clear every animation first, call
 * `clearSlideAnimations`.
 *
 * `effect` is any entry of the reference desktop app's Entrance, Emphasis and Exit
 * galleries (see `AnimationEffect`), written exactly as the reference desktop app writes it:
 * the same preset numbers and the same behaviours, so the deck plays the same
 * in the reference desktop app as an effect picked from its gallery.
 *
 * Effect Options are the options the reference desktop app offers for the preset:
 *
 *   - `direction` — `'flyIn'` / `'flyOut'` take any edge or corner of the
 *     slide (`'bottom'` by default); `'wipe…'` and `'peek…'` one of the four
 *     edges; `'strips…'` one of the four corners (`'bottomLeft'` by default).
 *   - `orientation` — blinds, random bars and checkerboard (`'horizontal'`)
 *     and split (`'vertical'`).
 *   - `inOut` — shape and split: close in on the centre (`'in'`) or open out
 *     of it (`'out'`). The reference desktop app's Shape exit defaults to `'out'`.
 *   - `shape` — `'circle'` (the default), `'box'`, `'diamond'` or `'plus'`.
 *   - `spokes` — wheel: 1 (the default), 2, 3, 4 or 8.
 *   - `spinDirection` and `spinDegrees` — spin: `'clockwise'` (the default)
 *     or `'counterclockwise'`, through any positive angle (360 by default).
 *   - `scaleDirection` and `scalePercent` — grow/shrink: `'both'` axes (the
 *     default), `'horizontal'` or `'vertical'`, to any positive size (150 %).
 *   - `transparencyPercent` — transparency: 0–100 (50 by default).
 *   - `color` — the colour emphasis effects (fill, font, line, brush and
 *     object colour, colour pulse, grow with colour): a theme slot or
 *     `#RRGGBB`, optionally with colour transforms. Accent 2 by default, and
 *     Background 1 for colour pulse.
 *
 * An option passed for an effect that does not take it is an error rather
 * than a no-op.
 *
 * `durationMs` defaults to the reference desktop app's default for the preset
 * (`defaultAnimationDurationMs`). An effect made of several behaviours — a
 * Bounce, a Teeter — is scaled as a whole, the way the reference desktop app's Duration box
 * scales it. `appear` and `disappear` are instantaneous and write no timed
 * behaviour; `transparency` and `boldReveal` hold until the slide ends and
 * take no duration.
 *
 * `start` decides where the effect lands. The default `'click'` gives it a
 * click stop of its own, so the viewer sees it on the next click.
 * `'withPrevious'` and `'afterPrevious'` join the stop the last effect is in,
 * running alongside it or once it has finished — both off the same click. As
 * the slide's first effect neither has a predecessor, so both run when the
 * slide appears. `delayMs` waits that long once the start condition is met.
 *
 * `'afterPrevious'` throws when the slide's timing does not say when the effect
 * before it ends — it runs indefinitely, states no duration, repeats, runs
 * letter by letter, or starts from another node rather than at a fixed offset.
 * The slide is left untouched; `'click'` and `'withPrevious'` need no such
 * measurement.
 *
 * `build` is the reference desktop app's Effect Options "Sequence". `'asOneObject'` (the
 * default) animates the shape with its text as one effect. `'allAtOnce'` and
 * `'byParagraph'` give every paragraph its own effect: all of them starting
 * together, or each on its own `start` — so the default `'click'` advances a
 * paragraph per click. They share a single `<p:bldP>`, which is how the reference desktop app
 * and Google Slides both present the build as one animation.
 * `getSlideAnimations` reports each paragraph as its own step, targeting a
 * paragraph range.
 *
 * Reading the result back, including the order and start condition of every
 * effect on the slide, is `getSlideAnimations`.
 */
declare const setShapeAnimation: (shape: SlideShapeData, opts: AnimationOptions) => void;
/**
 * Returns the animation effect bound to this shape via the slide's
 * `<p:timing>` tree, or `null` when the shape has none. A shape carrying
 * several effects reports the first one in document order. Presets outside
 * the `AnimationEffect` tokens are reported as `null` rather than guessed at.
 */
declare const getShapeAnimation: (shape: SlideShapeData) => AnimationEffect | null;
/**
 * Every animation effect on the slide, in document order — the click order a
 * viewer sees. Effects this library cannot author are still listed, with
 * `editable: false`, so a caller can show them instead of losing them.
 */
declare const getSlideAnimations: (slide: SlideData) => readonly SlideAnimationStep[];
/**
 * Returns every shape on the slide that has an authored animation
 * effect (i.e. `getShapeAnimation(shape)` is not `null`). Pair to
 * `slideHasAnimations`. Useful for audit reports — "which shapes on
 * this slide actually animate?" before exporting to a video pipeline
 * that doesn't honor the reference desktop app's timing tree.
 */
declare const findShapesWithAnimation: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/** Removes every animation on the slide, keeping its media play controls. */
declare const clearSlideAnimations: (slide: SlideData) => void;
//#endregion
//#region src/api/fn/slide-title.d.ts
/**
 * Returns the slide's title text, or `null` if neither a `title` nor
 * a `ctrTitle` placeholder is present.
 */
declare const getSlideTitle: (slide: SlideData) => string | null;
/**
 * Returns the slide's body text, or `null` if no `body` placeholder
 * is present. Mirror of `getSlideTitle`; pairs with `setSlideBody`.
 */
declare const getSlideBody: (slide: SlideData) => string | null;
/**
 * Sets the slide's title text. Looks for a `title` placeholder first,
 * falling back to `ctrTitle`. Throws if neither exists — the slide's
 * layout has no title slot.
 */
declare const setSlideTitle: (slide: SlideData, title: string) => void;
/**
 * Bulk-fills slide placeholders by type token. Each entry in
 * `byType` maps a `<p:ph type>` token (e.g. `'title'`, `'body'`,
 * `'ftr'`, `'dt'`) to the text to set. Silently skips entries
 * whose placeholder isn't present on the slide.
 *
 * Useful for template-fill workflows where the caller has all the
 * data in one struct.
 */
declare const setSlidePlaceholders: (slide: SlideData, byType: Readonly<Partial<Record<PlaceholderType, string>>>) => void;
/**
 * Writes `text` into the first body placeholder on the slide.
 * Newlines start a new paragraph (each becomes its own bullet on
 * layouts that bullet their body placeholder).
 *
 * Throws when the slide has no body placeholder — pair with
 * `findSlideLayoutByType(pres, 'obj')` / `'tx'` to add the slide
 * onto a layout that has one.
 */
declare const setSlideBody: (slide: SlideData, text: string) => void;
//#endregion
//#region src/api/fn/shape-runs.d.ts
/** Number of paragraphs; zero for an autoshape without a text body. Throws for non-text shapes. */
declare const getShapeParagraphCount: (shape: SlideShapeData) => number;
/**
 * One inline element in a paragraph as ordered: a literal text run
 * (`<a:r>`), a field substitution (`<a:fld>` — slide number, date, etc.),
 * or a line break (`<a:br>`). Renderers walk this list instead of the
 * strict `<a:r>`-only `getShapeRunCount` / `getShapeRunText` pair when
 * they need to reproduce the paragraph's full visible content.
 *
 * `text` is the cached value (`<a:t>` content for `r` and `fld`; `''`
 * for `br`). `format` is the literal `<a:rPr>` on the element when
 * present; use `getShapeRunFormatEffective` to walk inheritance.
 *
 * Field kinds (`fld.type`): typical ECMA-376 `ST_TextFieldType` tokens
 * are `slidenum`, `datetime` (variants `1`..`13`), `presentationDate`,
 * `headerfooter`, `footer`, etc. Unrecognised tokens come through
 * unchanged so renderers can decide whether to substitute live values.
 */
type ShapeParagraphElement = {
  readonly kind: 'r';
  readonly text: string;
  readonly format: ReadTextFormat | null;
} | {
  readonly kind: 'fld';
  readonly text: string;
  readonly format: ReadTextFormat | null;
  readonly type: string | null;
} | {
  readonly kind: 'br';
  readonly format: ReadTextFormat | null;
};
/**
 * Returns the inline children of a paragraph in document order — runs,
 * field placeholders, and line breaks. Used by renderers that need to
 * reproduce the paragraph faithfully (the `<a:r>`-only run accessors
 * silently drop fields and breaks). Omit the index to read all paragraphs in
 * one pass.
 */
declare function getShapeParagraphElements(shape: SlideShapeData): ReadonlyArray<ReadonlyArray<ShapeParagraphElement>>;
declare function getShapeParagraphElements(shape: SlideShapeData, paragraphIndex: number): ReadonlyArray<ShapeParagraphElement>;
/**
 * Reads the literal format of a paragraph's end mark (`<a:endParaRPr>`), or
 * `null` when the paragraph carries none. It is the only format a paragraph
 * with no runs carries. The write side is `setShapeParagraphs`' `endFormat`.
 */
declare const getParagraphEndFormat: (shape: SlideShapeData, paragraphIndex: number) => ReadTextFormat | null;
/**
 * Number of text runs in the given paragraph. Throws on out-of-range
 * paragraph index or non-text shapes.
 */
declare const getShapeRunCount: (shape: SlideShapeData, paragraphIndex: number) => number;
/** Visible text of a single run. */
declare const getShapeRunText: (shape: SlideShapeData, paragraphIndex: number, runIndex: number) => string;
/**
 * Sets `<a:hlinkClick>` on a single run. Per-run counterpart to
 * `setShapeHyperlink` (which targets every run in the shape). Pass
 * `null` to clear the link on that run alone — other runs are
 * untouched. Allocates or reuses a hyperlink rel on the slide
 * exactly like the shape-level setter.
 */
declare const setShapeRunHyperlink: (shape: SlideShapeData, paragraphIndex: number, runIndex: number, url: string | null, tooltip?: string) => void;
/**
 * Reads the external URL on a single run's `<a:hlinkClick>`. Per-run
 * counterpart to `getShapeHyperlink` (which only surfaces the first
 * link it finds). Returns `null` when this run has no link, or the
 * link's `r:id` resolves to a non-hyperlink rel.
 */
declare const getShapeRunHyperlink: (shape: SlideShapeData, paragraphIndex: number, runIndex: number) => string | null;
/**
 * Reads the tooltip text on the shape's `<a:hlinkClick tooltip="…"/>`.
 * Returns `null` when no hyperlink is set or the link doesn't author
 * a tooltip. Tooltips show up in the reference desktop app when the user hovers over
 * a linked shape in slide-show mode.
 *
 * Scans run-level `<a:rPr><a:hlinkClick>` first (where
 * `setShapeHyperlink` writes) and falls back to the
 * `<p:nvSpPr><p:cNvPr><a:hlinkClick>` shape-click hyperlink. Mirrors
 * `getShapeHyperlink`'s read path so the writer / reader pair is
 * consistent.
 */
declare const getShapeHyperlinkTooltip: (shape: SlideShapeData) => string | null;
/**
 * Reads the tooltip on a per-run `<a:rPr><a:hlinkClick tooltip="…"/>`.
 * Same semantics as `getShapeHyperlinkTooltip` but scoped to a single
 * run.
 */
declare const getShapeRunHyperlinkTooltip: (shape: SlideShapeData, paragraphIndex: number, runIndex: number) => string | null;
/**
 * Same as `getShapeClickAction` but reads the per-run
 * `<a:rPr><a:hlinkClick action=… r:id=…/>`. Recognises:
 *
 *   - `{ kind: 'url', url }` — external hyperlink rel
 *   - `{ kind: 'slide', slide }` — slide-jump action + slide rel
 *   - `{ kind: 'nextSlide' | 'prevSlide' | 'firstSlide' | 'lastSlide' }`
 *
 * Returns `null` for runs without an action or unknown action tokens.
 */
declare const getShapeRunClickAction: (shape: SlideShapeData, paragraphIndex: number, runIndex: number) => ShapeClickAction | null;
/**
 * Sets the horizontal alignment of a single paragraph. Same token set
 * as `setShapeAlignment`. Other paragraphs are untouched.
 */
declare const setParagraphAlignment: (shape: SlideShapeData | TableCellData, paragraphIndex: number, align: ParagraphAlignment) => void;
/**
 * Sets the paragraph's nesting level (`<a:pPr lvl="N"/>`). Levels are
 * 0-indexed; the reference desktop app accepts 0 through 8. Pass `0` to clear an
 * existing level — `<a:pPr lvl="0"/>` is the same as omitting the attr.
 * Indents matching the previous level's default bullet pair follow the level;
 * other indent values are preserved.
 *
 * A UTF-16 range selects all touched paragraphs (exclusive end; a caret selects
 * its paragraph). Pass `{ offset }` to shift their existing levels, clamped to
 * 0 through 8. Range updates preserve run XML and commit the text body once.
 *
 * Used in tandem with bullets to author nested lists:
 *
 *   setShapeText(shape, 'Item 1\nNested\nItem 2');
 *   setShapeBulletStyle(shape, 'bullet');
 *   setParagraphLevel(shape, 1, 1);  // indent the second line
 */
declare const setParagraphLevel: (shape: SlideShapeData | TableCellData, paragraphIndex: number | {
  start: number;
  end: number;
}, level: number | {
  offset: number;
}) => void;
/**
 * Reads the paragraph's own `algn` as its spec token (`l`, `ctr`, `r`,
 * `just`, `dist`, `justLow`, `thaiDist`) — so `setParagraphAlignment(…,
 * 'center')` reads back as `'ctr'`. Returns `null` when the attribute is
 * absent (the paragraph inherits from its layout / master) or is not a
 * valid token. For the inherited value under a plain-English name, use
 * `getParagraphPropertiesEffective`.
 */
declare const getParagraphAlignment: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => ParagraphAlignmentToken | null;
/**
 * Reads the paragraph's nesting level (`lvl` attribute), or `0` when
 * absent — the reference desktop app's default. A UTF-16 range returns the levels of all
 * touched paragraphs in order. Invalid indices or ranges throw RangeError.
 */
declare function getParagraphLevel(shape: SlideShapeData | TableCellData, paragraphIndex: number): number;
declare function getParagraphLevel(shape: SlideShapeData | TableCellData, range: {
  start: number;
  end: number;
}): number[];
/**
 * Sets the spacing before and/or after a paragraph, in points (where
 * a "point" is 1/72 inch). The reference desktop app stores these as hundredths of a
 * point inside `<a:pPr><a:spcBef>/<a:spcAft><a:spcPts val="…"/>` —
 * the helper converts.
 *
 *   setParagraphSpacing(shape, 0, { beforePts: 6, afterPts: 3 });
 *
 * Omitting a side keeps the existing value (or layout default).
 * Passing a side as `null` removes that spacing element.
 */
declare const setParagraphSpacing: (shape: SlideShapeData | TableCellData, paragraphIndex: number, opts: {
  beforePts?: number | null;
  afterPts?: number | null;
}) => void;
/**
 * Reads back paragraph spacing in points. Returns `{ beforePts,
 * afterPts }`; each side is `null` when no `<a:spcBef>` / `<a:spcAft>`
 * is present or when the inner element isn't `<a:spcPts>` (percentage
 * spacing is reported as `null` for now).
 */
declare const getParagraphSpacing: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => {
  readonly beforePts: number | null;
  readonly afterPts: number | null;
};
/**
 * Reads the paragraph's left / right / first-line indents from
 * `<a:pPr marL="…" marR="…" indent="…"/>`. Each is in EMU (matching
 * the reference desktop app's internal storage); positive means a positive indent,
 * negative `indent` is a hanging indent (typical for bullets).
 *
 * Returns `null` for sides the paragraph doesn't set (those inherit
 * from the layout / master).
 */
declare const getParagraphIndent: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => {
  leftEmu: number | null;
  rightEmu: number | null;
  firstLineEmu: number | null;
};
/**
 * Sets paragraph indents in EMU. The first-line offset is relative to the
 * left indent; negative values create a hanging indent. Omitted sides are
 * preserved, and `null` removes an override to restore inheritance.
 * Values are rounded to whole EMU and must satisfy the OOXML text bounds.
 */
declare const setParagraphIndent: (shape: SlideShapeData | TableCellData, paragraphIndex: number, opts: {
  leftEmu?: number | null;
  rightEmu?: number | null;
  firstLineEmu?: number | null;
}) => void;
/**
 * Reads the paragraph's `<a:lnSpc>` line spacing. The reference desktop app stores
 * line spacing two ways:
 *
 *   - Multiple of the natural line height — `<a:spcPct val="150000"/>`
 *     (= 1.5×). Returns `{ kind: 'pct', value }` with value as the unit
 *     fraction (1.5).
 *   - Fixed points — `<a:spcPts val="2400"/>` (= 24pt). Returns
 *     `{ kind: 'pts', value }` with value in points.
 *
 * Returns `null` when no `<a:lnSpc>` is present (the paragraph
 * inherits line spacing from the layout / master).
 */
declare const getParagraphLineSpacing: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => {
  readonly kind: "pct";
  readonly value: number;
} | {
  readonly kind: "pts";
  readonly value: number;
} | null;
/**
 * Sets a paragraph's line spacing — the writer counterpart to
 * `getParagraphLineSpacing`. Two modes (mirroring `<a:lnSpc>`):
 *
 *   - `{ kind: 'pct', value }` — a multiple of single spacing
 *     (`1` = single, `1.5` = 150%, `2` = double) → `<a:spcPct>`.
 *   - `{ kind: 'pts', value }` — a fixed leading in points → `<a:spcPts>`.
 *
 * Pass `null` to clear the override (the paragraph then inherits line
 * spacing from the layout / master).
 */
declare const setParagraphLineSpacing: (shape: SlideShapeData | TableCellData, paragraphIndex: number, spacing: {
  readonly kind: "pct";
  readonly value: number;
} | {
  readonly kind: "pts";
  readonly value: number;
} | null) => void;
/**
 * Reads back the bullet style on a single paragraph, or `null` when
 * no `<a:buChar>` / `<a:buAutoNum>` / `<a:buNone>` is present (the
 * paragraph inherits its bullet from the layout / master).
 */
declare const getParagraphBullet: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => BulletStyle | null;
/**
 * Returns `true` when the paragraph uses an image as its bullet
 * (`<a:pPr><a:buBlip r:embed="…"/>`). Renderers without image
 * support should fall back to a generic bullet glyph.
 *
 * The underlying rId / image bytes aren't surfaced here — resolving
 * that would need the rels of the layout / master the paragraph
 * inherits from, which can be cumbersome. Knowing that the bullet
 * *is* an image is usually enough for the UI to pick a fallback.
 */
declare const isParagraphBulletPicture: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => boolean;
/**
 * Returns the embedded image bytes backing a paragraph's picture bullet
 * (`<a:pPr><a:buBlip><a:blip r:embed="…"/>`), or `null` when the
 * paragraph has no picture bullet or the relationship can't be resolved
 * (external `r:link`, missing media part, etc.).
 *
 * The `r:embed` is resolved against the *slide's* relationships, matching
 * `getShapeImageBytes` — a bullet whose blip lives only on an inherited
 * layout/master paragraph isn't followed. The returned `Uint8Array` is a
 * live view into the package media part; treat it as read-only.
 */
declare const getParagraphBulletImageBytes: (shape: SlideShapeData | TableCellData, paragraphIndex: number) => Uint8Array | null;
/**
 * Reads the bullet's per-paragraph color, size, and font overrides —
 * `<a:buClr>` (theme-resolved hex), `<a:buSzPct>` / `<a:buSzPts>`
 * (size relative to run or fixed pt), and `<a:buFont typeface="…"/>`.
 *
 * Returns `{ color: null, sizePct: null, sizePts: null, font: null }`
 * when the paragraph doesn't override any of them (the bullet inherits
 * from the run / layout).
 */
declare const getParagraphBulletStyle: (pres: PresentationData, shape: SlideShapeData | TableCellData, paragraphIndex: number) => {
  color: string | null;
  sizePct: number | null;
  sizePts: number | null;
  font: string | null;
};
/**
 * Sets the bullet style on a single paragraph. Same `BulletStyle` shape
 * as `setShapeBulletStyle` — pass `'bullet'` / `'number'` / `'none'` or an
 * object like `{ char: '◆' }` / `{ autoNum: 'romanLcPeriod' }`.
 */
declare const setParagraphBullet: (shape: SlideShapeData | TableCellData, paragraphIndex: number, style: BulletStyle) => void;
/**
 * Sets the text of a single run. Existing rPr (font, size, color, ...)
 * is preserved — only the visible characters change.
 */
declare const setShapeRunText: (shape: SlideShapeData, paragraphIndex: number, runIndex: number, text: string) => void;
/**
 * Update paragraph line-breaking rules and character alignment. Omitted fields
 * are preserved; null removes the local override so the style can inherit.
 */
declare const setParagraphTypography: (shape: SlideShapeData | TableCellData, paragraphIndex: number, settings: {
  asianLineBreak?: boolean | null;
  latinLineBreak?: boolean | null;
  hangingPunctuation?: boolean | null;
  fontAlignment?: "auto" | "top" | "center" | "baseline" | "bottom" | null;
}) => void;
/** Set custom and automatic tab stops. Null restores inheritance; [] clears custom stops. */
declare const setParagraphTabs: (shape: SlideShapeData | TableCellData, paragraphIndex: number, settings: {
  tabStops?: readonly ParagraphTabStop[] | null;
  defaultTabSizeEmu?: number | null;
}) => void;
//#endregion
//#region src/api/fn/shape-removal-zorder.d.ts
/**
 * Copies a shape into `targetSlide`, including nested shapes and relationships.
 * Within a presentation, media and other referenced parts remain shared.
 * Across presentations, dependencies are cloned with collision-free part names,
 * retaining embedded workbooks, unknown parts and external links.
 * With `preserveGroupTransform`, a nested source retains its ancestor group
 * transforms as wrappers, excluding sibling objects. Returns the outermost
 * copied shape on the target slide.
 */
declare const copyShape: (targetSlide: SlideData, sourceShape: SlideShapeData, opts?: {
  readonly preserveGroupTransform?: boolean;
}) => SlideShapeData;
/** Move one shape or a selection in front of all siblings, preserving their stacking order. */
declare const bringShapeToFront: (shape: SlideShapeData | readonly SlideShapeData[]) => void;
/** Move one shape or a selection behind all siblings, preserving their stacking order. */
declare const sendShapeToBack: (shape: SlideShapeData | readonly SlideShapeData[]) => void;
/** Move one shape or a selection one step forward among siblings. */
declare const bringShapeForward: (shape: SlideShapeData | readonly SlideShapeData[]) => void;
/**
 * Returns the shape's z-index among its parent container's "real" shape children
 * (`<p:sp>` / `<p:pic>` / `<p:cxnSp>` / `<p:graphicFrame>` / `<p:grpSp>`),
 * skipping the required `<p:nvGrpSpPr>` / `<p:grpSpPr>` preface.
 * Higher numbers render in front.
 */
declare const getShapeZIndex: (shape: SlideShapeData) => number;
/**
 * Moves a shape or an ordered batch of siblings to a specific z-index.
 * A batch is inserted contiguously in the supplied order, back to front.
 * Index is clamped to the available range. Non-shape XML stays in place.
 */
declare const setShapeZIndex: (shape: SlideShapeData | readonly SlideShapeData[], toIndex: number) => void;
/** Move one shape or a selection one step backward among siblings. */
declare const sendShapeBackward: (shape: SlideShapeData | readonly SlideShapeData[]) => void;
/**
 * Removes every shape (sp / pic / cxnSp / graphicFrame / grpSp) from
 * the slide's `<p:spTree>`. The required `<p:nvGrpSpPr>` and
 * `<p:grpSpPr>` preface stays in place, so the slide is still valid
 * and re-applies its layout's placeholders on the next open.
 *
 * Useful for "start this slide over but keep its layout binding."
 */
declare const clearSlideShapes: (slide: SlideData) => void;
/**
 * Removes the shape from its slide's shape tree. Subsequent property
 * reads on this handle reflect the stale snapshot — discard it after.
 *
 * Removing a picture does NOT delete the underlying media part — it
 * may be referenced from other slides.
 */
declare const removeShape: (shape: SlideShapeData) => void;
//#endregion
//#region src/api/fn/shape-authoring.d.ts
/**
 * Adds a free-form text box to the slide. Returns the new shape.
 *
 * The box is a plain rectangle with no fill or outline carrying one
 * paragraph with one run. The shape id is allocated as one more than
 * the current max id.
 */
declare const addSlideTextBox: (slide: SlideData, opts: {
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  text: string;
  name?: string;
}) => SlideShapeData;
/**
 * Adds a preset shape (rectangle, ellipse, arrow, ...) to the slide.
 * Optional `text` seeds a single run.
 */
declare const addSlideShape: (slide: SlideData, opts: {
  preset: PresetShape | string;
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  text?: string;
  /** Vertical text anchor (`t` / `ctr` / `b`). Horizontal alignment is set
   * separately via `setShapeAlignment` / `setParagraphAlignment`. */
  textAnchor?: "ctr" | "t" | "b";
  name?: string;
}) => SlideShapeData;
/** Adds a straight-line connector between two points. */
declare const addSlideLine: (slide: SlideData, opts: {
  from: {
    x: Emu;
    y: Emu;
  };
  to: {
    x: Emu;
    y: Emu;
  };
  color?: Color;
  widthEmu?: number;
  name?: string;
}) => SlideShapeData;
/**
 * Adds a table to the slide. Cells render as plain text; `firstRow` /
 * `bandRow` flags drive the reference desktop app's banded-header look unless options say
 * otherwise.
 *
 * Cell text color is baked from the deck's resolved body-text color so the
 * table stays readable even on templates with an inverted color map
 * (`bg1="dk1" tx1="lt1"`), where the default `tx1` text token would otherwise
 * paint the text the same as the background. Override per cell afterwards with
 * `setTableCellTextFormat`.
 */
declare const addSlideTable: (slide: SlideData, opts: {
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  rows: ReadonlyArray<ReadonlyArray<string>>;
  colWidths?: ReadonlyArray<Emu>;
  rowHeights?: ReadonlyArray<Emu>;
  firstRow?: boolean;
  bandRow?: boolean;
  name?: string;
}) => SlideShapeData;
/**
 * Adds a picture to the slide from raw bytes. Returns the new shape.
 *
 * Allocates a `/ppt/media/imageN.<ext>` part, registers a Content_Types
 * Default if the extension isn't yet covered, allocates a slide→image
 * rel, and appends a `<p:pic>` element to the slide's `<p:spTree>`.
 *
 * Format is detected from magic bytes; pass `opts.format` to override.
 *
 * `opts.fit` controls how the image fills the `w × h` box. `'fill'` (the
 * default) stretches to the exact box, ignoring aspect ratio. `'contain'`
 * scales the image to fit inside the box preserving its aspect ratio and
 * centers it — measured from the PNG / JPEG header. For other formats (or
 * an unreadable header) `'contain'` falls back to `'fill'` rather than
 * erroring, since the natural size is unknown.
 */
declare const addSlideImage: (slide: SlideData, bytes: Uint8Array, opts: {
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  format?: ImageFormat;
  name?: string;
  fit?: ImageFit;
}) => SlideShapeData;
//#endregion
//#region src/api/fn/layout-edit.d.ts
/**
 * Renames the layout. This is the name the reference desktop app shows in the
 * "New Slide" gallery and in the layout pane, so it is what an author
 * picks a layout by — it is not an identifier, and nothing references
 * it.
 */
declare const setSlideLayoutName: (layout: SlideLayoutData, name: string) => void;
/**
 * Gives the layout a solid background, replacing whatever it had.
 * Every slide on this layout that does not set a background of its own
 * follows.
 *
 * `color` takes the same forms as `setSlideBackground`: `#RRGGBB` or
 * `scheme:accent1`.
 */
declare const setSlideLayoutBackground: (layout: SlideLayoutData, color: string) => void;
/**
 * Drops the layout's own background so it inherits the master's again.
 */
declare const clearSlideLayoutBackground: (layout: SlideLayoutData) => void;
/**
 * Moves and resizes one of the layout's placeholder slots, identified by
 * its position in `getSlideLayoutPlaceholders`.
 *
 * Slides inherit the new box wherever they have not pushed the
 * placeholder around themselves: a slide placeholder carrying its own
 * `<a:xfrm>` keeps it, exactly as the reference desktop app behaves when you edit a
 * layout under slides that were already nudged.
 */
declare const setSlideLayoutPlaceholderBounds: (layout: SlideLayoutData, index: number, bounds: ShapeBounds) => void;
//#endregion
//#region src/api/fn/_part-placeholders.d.ts
/**
 * Read-only view of one placeholder on a slide layout or master. Surfaces the
 * fields a slide-author cares about when binding a slide to a layout: which
 * slot is for the title, which is for the body, etc.
 */
interface SlideLayoutPlaceholder {
  /** `<p:ph type="...">`. `null` when omitted — the spec default is `obj`. */
  readonly type: string | null;
  /** `<p:ph idx="...">`. `null` when omitted — spec default is `0`. */
  readonly idx: number | null;
  /** `<p:cNvPr name="...">` — what the reference desktop app shows in the selection pane. */
  readonly name: string;
  /**
   * Position + size in EMU. A layout placeholder with no `<a:xfrm>` of its
   * own inherits them from the master. `null` when the placeholder has no
   * explicit transform.
   */
  readonly bounds: ShapeBounds | null;
}
//#endregion
//#region src/api/fn/layouts.d.ts
/** The reference desktop app's user-visible layout name. */
declare const getSlideLayoutName: (layout: SlideLayoutData) => string;
/**
 * Returns the package part name (e.g. `/ppt/slideLayouts/slideLayout3.xml`)
 * of `layout`. Useful for surfacing layouts in validator output and
 * other path-keyed UIs.
 */
declare const getSlideLayoutPartName: (layout: SlideLayoutData) => string;
/**
 * Returns the slide layout whose package part name equals
 * `partName`, or `null` when no such layout exists. Mirror of
 * `findSlideByPartName` for layouts.
 */
declare const findSlideLayoutByPartName: (pres: PresentationData, partName: string) => SlideLayoutData | null;
/**
 * Enumerates the placeholder shapes on a slide layout. Non-placeholder
 * shapes (decorative rectangles, watermarks added to the layout) are
 * filtered out; only entries with a `<p:ph>` element are returned.
 *
 * Use this when you need to discover which placeholder indices a
 * layout exposes — e.g. before `findSlidePlaceholder(slide, ...)` to
 * confirm the slot exists.
 */
declare const getSlideLayoutPlaceholders: (layout: SlideLayoutData) => ReadonlyArray<SlideLayoutPlaceholder>;
/**
 * Finds the first slide layout whose user-visible name matches `name`,
 * or `null` if none does. Accepts a literal string (exact equality)
 * or a `RegExp` for pattern matches — useful when template providers
 * suffix versions onto names (`'Title and Content v2'`).
 *
 * String matching is **case-sensitive** and exact: `findSlideLayout(pres,
 * 'Blank')` matches a layout named `"Blank"` but not `"blank"`. The
 * user-visible name is also locale-dependent (a deck authored in a
 * The reference desktop app localizes `"Blank"`), so prefer
 * {@link findSlideLayoutByType} — which matches the locale-stable
 * `<p:sldLayout type="…">` token (`'blank'`, `'title'`, `'obj'`, …) —
 * when you need a robust lookup. Pass a `RegExp` with the `i` flag here
 * for a case-insensitive name match.
 */
declare const findSlideLayout: (pres: PresentationData, name: string | RegExp) => SlideLayoutData | null;
/**
 * Returns every slide layout in the package that exposes a
 * placeholder of the given type token (`'title'`, `'body'`,
 * `'ftr'`, etc.). Useful for "find every layout that can host a
 * body" lookups before `addSlide`.
 */
declare const findLayoutsWithPlaceholderType: (pres: PresentationData, type: PlaceholderType) => ReadonlyArray<SlideLayoutData>;
/**
 * Finds the first slide layout with the given `<p:sldLayout type="...">`
 * token. Unlike `findSlideLayout` (which matches the user-visible
 * name, and is therefore locale-sensitive), this matches the spec
 * token — `title`, `obj`, `twoObj`, `blank`, etc. — and is stable
 * across the reference desktop app's UI languages.
 */
declare const findSlideLayoutByType: (pres: PresentationData, layoutType: SlideLayoutType | string) => SlideLayoutData | null;
/**
 * Layout type token, when present (`title`, `obj`, `twoObj`, ...).
 * `null` when omitted — the spec default for that case is `cust`.
 */
declare const getSlideLayoutType: (layout: SlideLayoutData) => SlideLayoutType | string | null;
/**
 * Enumerates every slide layout in the package.
 */
declare const getSlideLayouts: (pres: PresentationData) => ReadonlyArray<SlideLayoutData>;
/**
 * Layouts in the package that no slide references. Useful when
 * trimming a template deck — these layouts contribute parts and rels
 * without ever rendering. Iteration order matches `getSlideLayouts`.
 */
declare const getUnusedSlideLayouts: (pres: PresentationData) => ReadonlyArray<SlideLayoutData>;
/**
 * Returns a map of layout name → number of slides that reference it.
 * Every layout enumerated by `getSlideLayouts` appears as a key (zero
 * count for unreferenced layouts), so this surfaces unused layouts
 * directly — e.g. for templates that ship with placeholder layouts the
 * deck never picks up.
 */
/**
 * Histogram of slide-layout type token (e.g. `title`, `obj`, `twoObj`,
 * `blank`) → number of slides that use a layout of that type. Useful
 * for "how many content slides vs. dividers vs. title slides?" audits
 * — keyed on the OOXML enum, so stable across locales (unlike
 * `getSlideLayoutUsageCounts`, which keys on the user-visible name).
 * Layouts with no `type` token are skipped (the spec default is
 * `cust`).
 */
declare const getSlideLayoutUsageCountsByType: (pres: PresentationData) => Readonly<Record<string, number>>;
declare const getSlideLayoutUsageCounts: (pres: PresentationData) => Readonly<Record<string, number>>;
//#endregion
//#region src/api/fn/slide-masters.d.ts
/** The master's layouts, in the order the reference desktop app's Slide Master pane lists them. */
declare const getSlideMasterLayouts: (pres: PresentationData, master: string) => ReadonlyArray<SlideLayoutData>;
/**
 * The master's own placeholders (title, body, date, footer, slide number) and
 * where it puts them. A layout placeholder without a transform of its own sits
 * where the master's placeholder of the same type does.
 */
declare const getSlideMasterPlaceholders: (pres: PresentationData, master: string) => ReadonlyArray<SlideLayoutPlaceholder>;
/**
 * The master's name, as the reference desktop app shows it in Slide Master view ("<theme
 * name> Slide Master"). The reference desktop app keeps it as the name of the master's theme
 * (`<a:theme name>`), not in the master part, so renaming a master renames its
 * theme.
 */
declare const getSlideMasterName: (pres: PresentationData, master: string) => string;
/**
 * Whether the master is preserved (`<p:sldMaster preserve="1">`): the reference desktop app
 * deletes a master that is not preserved once no slide uses it.
 */
declare const isSlideMasterPreserved: (pres: PresentationData, master: string) => boolean;
/** Renames the master (its theme; see `getSlideMasterName`). */
declare const setSlideMasterName: (pres: PresentationData, master: string, name: string) => void;
/** Sets or clears Preserve Master. */
declare const setSlideMasterPreserved: (pres: PresentationData, master: string, preserved: boolean) => void;
/**
 * Insert Slide Master: adds the reference desktop app's default master, with its own
 * copy of the default theme named "Custom Design" and the eleven default
 * layouts, after the existing masters. The master is preserved, as the reference desktop app
 * marks a master it inserts. Returns the new master's part name.
 */
declare const addSlideMaster: (pres: PresentationData) => string;
/**
 * Deletes a master with its layouts and, when nothing else uses it, its theme.
 * Like the reference desktop app's Delete, this refuses (throws) while any slide uses one of
 * the master's layouts, and for the presentation's only master.
 */
declare const removeSlideMaster: (pres: PresentationData, master: string) => void;
/**
 * Insert Layout: adds the reference desktop app's "Custom Layout" — a title and the date,
 * footer and slide-number placeholders — to `master`, at `index` in its layout
 * list (the end by default). Repeats are named "1_Custom Layout", and so on.
 */
declare const addSlideLayout: (pres: PresentationData, master: string, options?: {
  index?: number;
}) => SlideLayoutData;
/**
 * Deletes a layout. Like the reference desktop app's Delete, this refuses (throws) while any
 * slide uses the layout, and for a master's only layout.
 */
declare const removeSlideLayout: (layout: SlideLayoutData) => void;
/**
 * Master Layout: removes one of the master's five placeholders, or puts it
 * back where the reference desktop app's default master has it.
 */
declare const setSlideMasterPlaceholderIncluded: (pres: PresentationData, master: string, type: MasterPlaceholderType, included: boolean) => void;
/** The layout's Title checkbox: removes its title, or adds one that follows the master's. */
declare const setSlideLayoutTitleIncluded: (layout: SlideLayoutData, included: boolean) => void;
/**
 * The layout's Footers checkbox: removes its date, footer and slide-number
 * placeholders, or adds whichever are missing (they follow the master's).
 */
declare const setSlideLayoutFootersIncluded: (layout: SlideLayoutData, included: boolean) => void;
/**
 * Hide Background Graphics on a layout (`showMasterSp="0"`): slides on it stop
 * showing the master's decorative shapes; the layout's own shapes stay.
 */
declare const setSlideLayoutBackgroundGraphicsHidden: (layout: SlideLayoutData, hidden: boolean) => void;
/**
 * Insert Placeholder: draws a placeholder of `kind` on the layout at `bounds`.
 * Slides added on the layout get the placeholder; content and text
 * placeholders carry the master's five prompt levels, the others their
 * kind's prompt ("Picture", "Chart", …). Returns the placeholder's index in
 * `getSlideLayoutPlaceholders`.
 */
declare const addSlideLayoutPlaceholder: (layout: SlideLayoutData, kind: LayoutPlaceholderKind, bounds: ShapeBounds) => number;
//#endregion
//#region src/api/fn/page-masters.d.ts
/**
 * The notes master's placeholders (header, date, slide image, notes body,
 * footer, slide number) and where they sit on the page, or `null` when the
 * presentation has no notes master.
 */
declare const getNotesMasterPlaceholders: (pres: PresentationData) => ReadonlyArray<SlideLayoutPlaceholder> | null;
/**
 * The handout master's placeholders (header, date, footer, page number), or
 * `null` when the presentation has no handout master.
 */
declare const getHandoutMasterPlaceholders: (pres: PresentationData) => ReadonlyArray<SlideLayoutPlaceholder> | null;
/**
 * The Notes Master tab's placeholder checkboxes: removes a placeholder from
 * the notes master, or puts it back where the reference desktop app's default puts it.
 * Creates the default notes master when the presentation has none.
 */
declare const setNotesMasterPlaceholderIncluded: (pres: PresentationData, type: NotesMasterPlaceholderType, included: boolean) => void;
/** The Handout Master tab's placeholder checkboxes; see `setNotesMasterPlaceholderIncluded`. */
declare const setHandoutMasterPlaceholderIncluded: (pres: PresentationData, type: HandoutMasterPlaceholderType, included: boolean) => void;
/**
 * The notes and handout page (`<p:notesSz>`), in EMU. Notes pages, handouts
 * and the outline print on it; it is portrait unless it is wider than tall.
 */
declare const getNotesPageSize: (pres: PresentationData) => {
  width: Emu;
  height: Emu;
};
/**
 * Handout Orientation / Notes Page Orientation. Both turn the one notes page
 * (`<p:notesSz>`) that notes pages and handouts share, and the reference desktop app moves
 * the notes master, the handout master and every notes page with it.
 */
declare const setNotesPageOrientation: (pres: PresentationData, orientation: "portrait" | "landscape") => void;
/** The Handout Master tab's Slides Per Page choices (`'outline'` is Slide Outline). */
type HandoutSlidesPerPage = 1 | 2 | 3 | 4 | 6 | 9 | 'outline';
/**
 * How many slides the handout master lays out per page. The handout master
 * part has nowhere to keep it; the schema's only home for a handout layout is
 * the print settings' `<p:prnPr prnWhat>` in the presentation properties, so
 * that is where it is read from (six when it names no handout layout).
 */
declare const getHandoutSlidesPerPage: (pres: PresentationData) => HandoutSlidesPerPage;
/** Sets the handout's slides per page; see `getHandoutSlidesPerPage`. */
declare const setHandoutSlidesPerPage: (pres: PresentationData, slidesPerPage: HandoutSlidesPerPage) => void;
//#endregion
//#region src/api/fn/properties.d.ts
/**
 * Document-level metadata from `/docProps/core.xml` (Open Packaging
 * Conventions). Surfaces the fields the reference desktop app, Keynote, and
 * everyone else exchange via OPC core-properties — these are the
 * values shown in the reference desktop app's "File › Properties" / "Info" panel.
 */
interface CoreProperties {
  readonly title: string | null;
  readonly subject: string | null;
  readonly creator: string | null;
  readonly keywords: string | null;
  readonly description: string | null;
  readonly lastModifiedBy: string | null;
  readonly revision: string | null;
  /** ISO-8601 timestamp string when set; `null` otherwise. */
  readonly created: string | null;
  /** ISO-8601 timestamp string when set; `null` otherwise. */
  readonly modified: string | null;
  readonly category: string | null;
}
/**
 * Reads `/docProps/core.xml`. Returns `null` when the package has
 * no core-properties part. Each field is `null` when the
 * corresponding element is absent or empty.
 */
/**
 * Convenience: bumps core-properties' `cp:revision` by one (treating
 * an unset / unparseable value as 0). Returns the new revision
 * number. Useful right before `savePresentation` so consumers can
 * tell decks apart.
 */
declare const incrementRevision: (pres: PresentationData) => number;
/**
 * Convenience: writes `new Date().toISOString()` to
 * `dcterms:modified`. Useful right before `savePresentation` so
 * "last edited" shows the actual save time. Pass an explicit
 * `Date` to set a specific value.
 */
declare const touchModified: (pres: PresentationData, at?: Date) => void;
/**
 * Convenience: the timestamp from core-properties' `dcterms:created`,
 * parsed as a `Date`. Returns `null` when no created field is set
 * or the value isn't a recognizable W3C-DTF / ISO-8601 string.
 */
declare const getPresentationCreated: (pres: PresentationData) => Date | null;
/**
 * Convenience: the timestamp from core-properties' `dcterms:modified`,
 * parsed as a `Date`. Returns `null` when no modified field is set
 * or the value isn't a recognizable W3C-DTF / ISO-8601 string.
 */
declare const getPresentationModified: (pres: PresentationData) => Date | null;
declare const getCoreProperties: (pres: PresentationData) => CoreProperties | null;
/**
 * Writes selected fields on `/docProps/core.xml`. Unspecified fields
 * are left as-is; pass `null` to clear a field that's currently set.
 * Bootstraps the part (and the `/_rels/.rels` entry + content-type
 * override) if the package didn't have one.
 *
 * Note: setting `created` / `modified` requires an ISO-8601 timestamp
 * string (e.g. `'2026-05-15T12:34:56Z'`). The reference desktop app expects the
 * `xsi:type="dcterms:W3CDTF"` attribute on these elements but readers
 * we tested all accept missing-attribute output too; this helper
 * therefore omits the attribute for simplicity.
 */
declare const setCoreProperties: (pres: PresentationData, values: Partial<CoreProperties>) => void;
/**
 * Selected string fields from `/docProps/app.xml`
 * (extended-properties / "app props"). The reference desktop app exposes these
 * under File › Info / Properties as the "Origin" and "Related
 * People" groups.
 *
 * Numeric / derived fields (`Slides`, `Words`, `Paragraphs`, …) are
 * intentionally omitted — they're recomputed by the reference desktop app on save
 * and reading them tends to lie about decks edited in other apps.
 */
interface ExtendedProperties {
  readonly application: string | null;
  readonly appVersion: string | null;
  readonly company: string | null;
  readonly manager: string | null;
  readonly presentationFormat: string | null;
  readonly hyperlinkBase: string | null;
}
/**
 * Reads `/docProps/app.xml`. Returns `null` if the package has no
 * extended-properties part. Each field is `null` when the
 * corresponding element is absent or empty.
 */
declare const getExtendedProperties: (pres: PresentationData) => ExtendedProperties | null;
/**
 * Writes selected fields on `/docProps/app.xml`. Throws when the
 * package has no extended-properties part — unlike core-properties,
 * we don't bootstrap app.xml from scratch because its schema
 * requires several derived `<vt:*>` elements (`HeadingPairs`,
 * `TitlesOfParts`, …) that aren't user-facing.
 *
 * Pass `null` to clear an existing field's text. Unspecified keys
 * are left untouched.
 */
declare const setExtendedProperties: (pres: PresentationData, values: Partial<ExtendedProperties>) => void;
//#endregion
//#region src/api/fn/thumbnail.d.ts
/**
 * The package's thumbnail image, when present. The reference desktop app, the OS
 * file picker, and SharePoint preview decks via this. Format is
 * what's encoded in the thumbnail part — usually JPEG.
 */
interface PresentationThumbnail {
  readonly format: ImageFormat;
  readonly bytes: Uint8Array;
}
/**
 * Returns the package's thumbnail bytes, plus the detected image
 * format. Returns `null` when the package has no thumbnail rel or
 * the rel target part is missing.
 *
 * The returned `bytes` is a live view into the thumbnail part —
 * treat it as read-only; copy if you need an independent buffer.
 */
declare const getThumbnail: (pres: PresentationData) => PresentationThumbnail | null;
/**
 * Replaces the package's thumbnail. Auto-detects the image format
 * from the bytes (pass `options.format` to override). Bootstraps the
 * thumbnail part + root rel + content-type default if the package
 * had no thumbnail; otherwise replaces the existing thumbnail in
 * place, switching its filename / extension if the format changed.
 */
declare const setThumbnail: (pres: PresentationData, bytes: Uint8Array, options?: {
  format?: ImageFormat;
}) => void;
/**
 * Removes the package's thumbnail entirely (drops the rel + the
 * underlying part). No-op when the package has no thumbnail.
 */
declare const removeThumbnail: (pres: PresentationData) => void;
//#endregion
//#region src/api/fn/slide-query.d.ts
/**
 * Enumerates slides in presentation order. Returns opaque `SlideData`
 * handles that pass to every slide-level fn-API helper.
 *
 * Throws if any referenced slide part is missing — a structurally
 * invalid PPTX cannot honor the L1 contract.
 */
/**
 * Returns the shape at the given 0-based `index` on the slide, or
 * `null` when `index` is out of range. Convenience over
 * `getSlideShapes(slide)[index] ?? null`.
 */
declare const getShapeAt: (slide: SlideData, index: number) => SlideShapeData | null;
/**
 * Returns the slide count without forcing every slide part to be
 * parsed. Reads only `presentation.xml` and walks `<p:sldIdLst>`.
 *
 * Equivalent to `getSlides(pres).length`, but cheaper on cold reads
 * — useful when the caller only wants a count for UI badges or
 * validation logic.
 */
declare const getSlideCount: (pres: PresentationData) => number;
/**
 * Returns the number of slide layouts in the package. Walks the
 * content-type map; cheaper than `getSlideLayouts(pres).length`
 * because it avoids parsing each layout's XML body.
 */
declare const getSlideLayoutCount: (pres: PresentationData) => number;
declare const getSlides: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Concatenated visible text of a slide. Convenience wrapper that walks
 * the slide's shape tree without instantiating any class.
 */
declare const getSlideText: (slide: SlideData) => string;
/**
 * Length (in code points) of the slide's concatenated visible text.
 * Counts Unicode code points via `Array.from`, so surrogate-pair
 * characters (emoji, supplementary CJK) count as 1, matching the
 * library's invariant on text editing.
 */
declare const getSlideTextLength: (slide: SlideData) => number;
/**
 * A single slide's outline entry: its index, title (or `null`), and
 * the text of its body placeholder (or `null`). Used by
 * `getSlideOutline`.
 */
interface SlideOutlineEntry {
  readonly index: number;
  readonly title: string | null;
  readonly body: string | null;
}
/**
 * Returns a one-entry-per-slide outline: title + body-placeholder
 * text. Useful for building thumbnail panes, table-of-contents
 * inserts, or quick-look previews without rendering the slide.
 *
 * "Body" here is the first matching `body` placeholder per
 * `findSlidePlaceholder(slide, 'body')`. Non-placeholder text on
 * the slide isn't surfaced; use `getSlideText(slide)` for that.
 */
declare const getSlideOutline: (pres: PresentationData) => ReadonlyArray<SlideOutlineEntry>;
/**
 * Concatenated title + body text from every slide, joined with the
 * given `separator` (defaults to `'\n\n'`). Useful for generating
 * a table-of-contents handout from a deck. Slides without a title
 * still contribute their body (and vice versa); slides without
 * both are skipped.
 */
declare const getOutlineText: (pres: PresentationData, separator?: string) => string;
/**
 * Concatenated visible text from every slide, joined with the
 * given `separator` (defaults to a form-feed, `\f`, between slides).
 * Useful for search-indexing a whole deck without iterating slides
 * yourself.
 */
declare const getPresentationText: (pres: PresentationData, separator?: string) => string;
/**
 * Total code-point length of visible text across every slide.
 * Sibling of `getSlideTextLength`; counts surrogate-pair characters
 * (emoji, supplementary CJK) as 1 each. Cheaper than building the
 * concatenated string when only the length matters.
 */
declare const getPresentationTextLength: (pres: PresentationData) => number;
/**
 * Dense per-slide shape count array, 0-based by slide index. Counts
 * top-level + group-children shapes (whatever `getSlideShapes` flattens).
 * Useful for charting shape density per slide and identifying outliers
 * (the 200-shape "soup" slide everybody complains about) for cleanup.
 */
declare const getPresentationShapeCountsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * Dense per-slide visible-text length array, 0-based by slide index.
 * Counts code points (surrogate pairs as 1) per `getSlideTextLength`.
 * Pair with `getPresentationShapeCountsBySlide` for slide-density
 * audits — high text length on a slide with few shapes usually means
 * one paragraph-heavy text box; low text on a many-shape slide
 * usually means a busy diagram.
 */
declare const getPresentationTextLengthsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * Returns the 0-based index of `slide` within `pres`, or `-1` if the
 * slide doesn't belong to this presentation (e.g. after a removeSlide
 * call, or if it was constructed from a different package).
 */
declare const getSlideIndex: (pres: PresentationData, slide: SlideData) => number;
/**
 * Returns the slide at the given 0-based `index`, or `null` if `index`
 * is out of range. Convenience over `getSlides(pres)[index] ?? null`.
 */
declare const getSlideAt: (pres: PresentationData, index: number) => SlideData | null;
/**
 * Per-slide summary suitable for slide pickers or audit reports.
 *
 * `index` is the 0-based position in the deck. `title` falls back to
 * `null` when the slide has no title placeholder. `layoutName` falls
 * back to `null` when the slide has no `slideLayout` rel or the
 * layout doesn't carry a user-visible name.
 */
interface SlideInfo {
  readonly index: number;
  readonly title: string | null;
  readonly hidden: boolean;
  readonly shapeCount: number;
  readonly layoutName: string | null;
}
declare const getSlideInfo: (pres: PresentationData, slide: SlideData) => SlideInfo;
/**
 * Returns every slide bound to the given `layout`. Useful for "find
 * all slides using the `Title and Content` layout" workflows; pair
 * with `findSlideLayout(pres, name)` to look up the layout by name.
 */
declare const getSlidesByLayout: (pres: PresentationData, layout: SlideLayoutData) => ReadonlyArray<SlideData>;
/**
 * Returns the first slide whose title-placeholder text equals
 * `title` exactly. Returns `null` when no slide matches. Different
 * from `findSlideByText`, which matches any visible text — this one
 * is title-placeholder-scoped. Accepts a literal string (exact
 * equality) or a `RegExp` for pattern matches.
 */
declare const findSlideByTitle: (pres: PresentationData, title: string | RegExp) => SlideData | null;
/**
 * Returns the slide whose package part name equals `partName`
 * (typically `/ppt/slides/slideN.xml`), or `null` when no such
 * slide exists. Useful for callers that walk validator output,
 * `listPackageParts`, or any other low-level path-keyed API and
 * need the matching `SlideData` handle.
 */
declare const findSlideByPartName: (pres: PresentationData, partName: string) => SlideData | null;
/**
 * Returns the package part name (e.g. `/ppt/slides/slide3.xml`) of
 * `slide`. Useful for surfacing slides in validator output, error
 * messages, or any path-keyed sidebar UI.
 */
declare const getSlidePartName: (slide: SlideData) => string;
/**
 * Returns the slide's current XML body as a string. Re-serializes
 * from the typed AST, so this reflects any pending edits that
 * haven't been written back to the package yet.
 *
 * Intended for diagnostics: dumping into bug reports, asserting
 * structure in tests, snapshotting before / after a transformation.
 * Do NOT parse this back yourself — round-trip safely through
 * `loadPresentation` / `savePresentation`.
 */
declare const getSlideXmlString: (slide: SlideData) => string;
/**
 * Returns the first slide whose concatenated visible text contains
 * `needle` (substring; case-sensitive). Pass a `RegExp` to test
 * against the slide's text body instead. Returns `null` when no
 * slide matches.
 */
declare const findSlideByText: (pres: PresentationData, needle: string | RegExp) => SlideData | null;
/**
 * Every slide whose concatenated visible text contains `needle`
 * (substring) or matches the given `RegExp`.
 */
declare const findSlidesByText: (pres: PresentationData, needle: string | RegExp) => ReadonlyArray<SlideData>;
/**
 * Every slide whose speaker notes contain `needle` (substring) or
 * match the given `RegExp`. Sibling of `findSlidesByText` — useful
 * for surfacing reviewer-facing annotations stored in notes (e.g.
 * "every slide whose notes say TODO").
 *
 * Slides without notes are skipped.
 */
declare const findSlidesByNotes: (pres: PresentationData, needle: string | RegExp) => ReadonlyArray<SlideData>;
/**
 * Returns every slide where `needle` appears in **either** the
 * visible text or the speaker notes. Each slide is reported once,
 * in document order. Sibling of `findSlidesByText` /
 * `findSlidesByNotes` — useful for a generic "search the deck"
 * UX where the caller doesn't care which surface matched.
 */
declare const searchSlides: (pres: PresentationData, needle: string | RegExp) => ReadonlyArray<SlideData>;
/**
 * Replaces every occurrence of `from` in every slide's speaker
 * notes with `to`. Returns the number of slides updated. Sibling
 * of `replaceTextInPresentation` for the reviewer-annotation pass.
 *
 * Slides without notes are skipped. For string needles, replacement
 * is global; for RegExp needles, the global flag is forced.
 */
declare const replaceTextInNotes: (pres: PresentationData, from: string | RegExp, to: string) => number;
/**
 * Slide-scoped sibling of `replaceTextInNotes`. Replaces every
 * occurrence of `from` in the slide's speaker notes with `to`.
 * Returns `true` if the notes changed, `false` otherwise (no notes
 * present, or no match).
 */
declare const replaceTextInSlideNotes: (slide: SlideData, from: string | RegExp, to: string) => boolean;
/**
 * Every shape on every slide, paired with its slide and the slide's
 * 0-based index. Useful for cross-deck audits ("find every picture",
 * "list shape names anywhere") without writing the nested-loop
 * boilerplate yourself.
 */
interface AllShapesEntry {
  readonly slide: SlideData;
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
}
declare const getAllShapes: (pres: PresentationData) => ReadonlyArray<AllShapesEntry>;
/**
 * Returns `true` when the slide carries `show="0"` on its root
 * `<p:sld>` element. The reference desktop app hides such slides from the slideshow
 * but keeps them in the editing surface.
 */
declare const isSlideHidden: (slide: SlideData) => boolean;
/** Whether inherited master and layout decoration is hidden on this slide. */
declare const isSlideBackgroundGraphicsHidden: (slide: SlideData) => boolean;
/** Whether a layout suppresses its master's decoration, keeping its own shapes. */
declare const isSlideLayoutBackgroundGraphicsHidden: (layout: SlideLayoutData) => boolean;
/**
 * Hide inherited decoration without removing shapes or changing background fills.
 * The reference desktop app's Hide Background Graphics suppresses layout decoration as well as
 * master shapes, despite the attribute's name. Slide placeholders remain visible.
 */
declare const setSlideBackgroundGraphicsHidden: (slide: SlideData, hidden: boolean) => void;
/**
 * Returns `true` when the slide has at least one authored animation effect.
 * Per-slide complement to `getPresentationSummary().hasAnimations`, which only
 * reports a deck-wide flag.
 *
 * A slide holding only a video / audio clip also carries a `<p:timing>` block
 * — that is where the clip's play controls live — so the element's presence
 * alone is not an animation.
 */
declare const slideHasAnimations: (slide: SlideData) => boolean;
/**
 * Toggles the slide's visibility in the slideshow. Hiding adds
 * `show="0"`; showing removes the attribute (the reference desktop app treats absence
 * as the default `show="1"`).
 */
declare const setSlideHidden: (slide: SlideData, hidden: boolean) => void;
/**
 * Replaces `{{key}}` tokens on every slide. Returns the total number of
 * substitutions performed.
 */
declare const replaceTokensInPresentation: (pres: PresentationData, tokens: Record<string, string>) => number;
/**
 * Replaces every occurrence of `from` in every slide's text with `to`.
 * `from` may be a string (treated as a literal) or a `RegExp`. Returns
 * the number of `<a:t>` elements mutated across the whole deck. Matches can
 * span formatting runs within a paragraph, but not explicit line breaks.
 * Inserted text inherits the first matched run’s formatting. `to` follows
 * JavaScript replacement-string rules (`$$`, `$&`, capture groups).
 *
 * Use this for the broad "rename product X to Y" pattern; for
 * `{{token}}` style substitutions, prefer
 * `replaceTokensInPresentation` which is more explicit about intent.
 */
declare const replaceTextInPresentation: (pres: PresentationData, from: string | RegExp, to: string) => number;
/**
 * Replaces every occurrence of `from` in the slide's text with `to`.
 * Matches span adjacent formatting runs, bounded by paragraphs and explicit
 * line breaks. Inserted text inherits the first matched run’s formatting;
 * `to` follows JavaScript replacement-string rules. Returns the number of
 * `<a:t>` elements mutated on this slide.
 */
declare const replaceTextInSlide: (slide: SlideData, from: string | RegExp, to: string) => number;
//#endregion
//#region src/api/fn/custom-shows.d.ts
/** A named slide sequence. Repeated slides and empty sequences are supported. */
interface CustomShow {
  readonly id: number;
  readonly name: string;
  readonly slides: readonly SlideData[];
}
/** Reads custom shows in list order. Missing slide references are omitted. */
declare function getCustomShows(pres: PresentationData): CustomShow[];
/**
 * Replaces the custom-show list atomically, retaining metadata on surviving IDs.
 * Slide handles must belong to this presentation. IDs are stable unsigned integers.
 */
declare function setCustomShows(pres: PresentationData, shows: readonly CustomShow[]): void;
//#endregion
//#region src/api/fn/slide-deck.d.ts
/**
 * Convenience over `addSlide` that picks a layout automatically:
 *
 *   1. The layout with `<p:sldLayout type="blank">`, if present.
 *   2. Otherwise, the first available layout (alphabetical by
 *      part name).
 *
 * Throws when the package carries no layouts at all (which would
 * be a structurally-broken deck).
 */
declare const addBlankSlide: (pres: PresentationData) => SlideData;
/**
 * Sugar over `addSlide` + `setSlideTitle` + `setSlideBody` for the
 * "title + body" pattern. Picks the `obj` (Title and Content)
 * layout when present, falling back to the first layout with a
 * body placeholder.
 *
 * Throws if no layout in the package offers a body slot.
 */
declare const addContentSlide: (pres: PresentationData, opts: {
  title?: string;
  body?: string;
}) => SlideData;
/**
 * Sugar over `addSlide` + `setSlideTitle` for the section-divider
 * pattern. Picks `<p:sldLayout type="secHead">` when present (the
 * the reference desktop app's "Section Header" layout); otherwise falls back to a
 * `title`-typed layout or the first available layout.
 */
declare const addSectionHeaderSlide: (pres: PresentationData, title: string) => SlideData;
/**
 * Sugar over `addSlide` + `setSlideTitle` for the common
 * "title slide + set heading" pattern. Picks the `title` layout
 * first, then falls back to the first non-blank layout.
 *
 * Throws when the package carries no layouts at all.
 */
declare const addTitleSlide: (pres: PresentationData, title: string) => SlideData;
/**
 * Adds a new slide bound to `layout`. Returns the new `SlideData`.
 *
 * Allocates a fresh part name, sldId, and rId; clones layout
 * placeholders into the slide; writes `[Content_Types].xml`, the
 * slide's `.rels`, presentation's `.rels`, and `<p:sldIdLst>`. The
 * deck cache is updated while existing slide handles remain valid.
 * The layout must belong to this presentation.
 */
declare const addSlide: (pres: PresentationData, options: {
  layout: SlideLayoutData;
}) => SlideData;
/**
 * Removes the given slide from the deck. Removes the `<p:sldId>`, the
 * `presentation.xml.rels` entry, and the slide part + its `.rels` part.
 * Links to this slide from other slides are cleared to prevent accidental
 * retargeting when a slide part name is reused.
 *
 * Media parts are intentionally NOT cleaned up — they may be shared
 * with other slides. The freed `sldId` is NOT reused on subsequent
 * `addSlide` calls (the reference desktop app's quirk, see plan §Risks).
 */
declare const removeSlide: (pres: PresentationData, slide: SlideData) => void;
/**
 * Reorders every slide in the deck via a custom comparator. The
 * comparator is invoked with two `SlideData` handles and returns the
 * usual `Array.prototype.sort` ordering (-1 / 0 / 1).
 *
 *   sortSlides(pres, (a, b) => getSlideTitle(a)?.localeCompare(getSlideTitle(b) ?? '') ?? 0);
 *
 * Internally walks `<p:sldIdLst>` and re-emits its `<p:sldId>` children
 * in the new order. Slide parts and rels are untouched — only the
 * order in which the reference desktop app plays them changes.
 */
declare const sortSlides: (pres: PresentationData, compareFn: (a: SlideData, b: SlideData) => number) => void;
/**
 * Reverses the slide order across the whole deck. Built on
 * `sortSlides` for predictable rels behavior.
 */
declare const reverseSlides: (pres: PresentationData) => void;
/**
 * Swaps the positions of the slides at `indexA` and `indexB`.
 * No-op when the indices are equal. Throws on out-of-range indices.
 * Implemented on top of `moveSlide` for predictable rels behavior.
 */
declare const swapSlides: (pres: PresentationData, indexA: number, indexB: number) => void;
/**
 * Reorders a slide. The slide's identity (part, rels, sldId) is
 * unchanged — only `<p:sldIdLst>`'s child order changes.
 */
declare const moveSlide: (pres: PresentationData, slide: SlideData, toIndex: number) => void;
/**
 * Duplicates a slide. Returns the new `SlideData` appended to deck order.
 *
 * Charts, embedded workbooks, notes and other owned dependencies are copied.
 * Layouts, masters, themes, media and links to other slides remain shared.
 * Unknown dependency bodies are retained without interpreting their content.
 */
declare const duplicateSlide: (pres: PresentationData, slide: SlideData) => SlideData;
/**
 * Inserts slides at a 0-based index, clamped to the presentation bounds.
 * A list of layout options inserts a contiguous batch in the same order,
 * allocating IDs and updating shared package parts once. Existing slide handles
 * remain valid. All layouts must belong to the target presentation.
 */
declare function addSlideAt(pres: PresentationData, atIndex: number, options: ReadonlyArray<{
  layout: SlideLayoutData;
}>): SlideData[];
declare function addSlideAt(pres: PresentationData, atIndex: number, options: {
  layout: SlideLayoutData;
}): SlideData;
/**
 * Convenience over `duplicateSlide` + `moveSlide`. Duplicates `slide`
 * and inserts the duplicate at `atIndex` instead of at the end.
 */
declare const duplicateSlideAt: (pres: PresentationData, atIndex: number, slide: SlideData) => SlideData;
/**
 * Imports a slide into `targetPres`. When `targetLayout` is supplied, rebinds
 * the slide to it. Otherwise preserves the source layout, master and theme.
 * Copies the complete relationship graph, including charts, workbooks, notes,
 * media and unknown extension parts. Shared dependencies and cycles retain
 * their relationships; external hyperlinks are preserved. Missing dependencies
 * fail before writing any imported parts.
 * Returns the new slide appended to the target presentation.
 */
declare const importSlide: (targetPres: PresentationData, sourceSlide: SlideData, targetLayout?: SlideLayoutData) => SlideData;
/**
 * Appends every slide from `sourcePres` into `targetPres`, in source
 * order. Built on top of `importSlide`: dependencies are preserved, and the slide's
 * layout is rebound to `targetLayout` on the target side.
 *
 * `targetLayout` can be a single layout used for every imported
 * slide (common), or a function called once per source slide for
 * per-slide layout selection.
 *
 * Returns the imported slides in target order.
 */
declare const mergePresentations: (targetPres: PresentationData, sourcePres: PresentationData, targetLayout: SlideLayoutData | ((sourceSlide: SlideData, index: number) => SlideLayoutData)) => ReadonlyArray<SlideData>;
//#endregion
//#region src/api/fn/slide-placeholder-geometry.d.ts
/**
 * Restore top-level placeholders' position, size, rotation and flips from the
 * current layout (or its master when the layout inherits its geometry).
 * Text, formatting, IDs, relationships and other shapes are preserved. Slots
 * without matching layout geometry are left intact; deleted placeholders are
 * not recreated.
 *
 * A placeholder inside a group is left where it is. The layout states a
 * rectangle on the slide, while a grouped shape's own geometry is written in
 * its group's coordinate space — so restoring it would mean either tearing the
 * shape out of the arrangement it was grouped into, or inventing a rectangle
 * the layout never described. Its formatting and appearance are still reset by
 * the two functions below.
 *
 * Returns the number of placeholders restored.
 */
declare const resetSlidePlaceholderGeometry: (slide: SlideData) => number;
/**
 * Restore missing slide placeholders using the current layout.
 * Existing content, geometry and formatting are untouched, including grouped
 * placeholders. New slots inherit their layout's geometry and text style;
 * layout prompt text and layout-only relationships are not copied.
 * Returns the number of added placeholders. Repeated calls do not add duplicates.
 */
declare const addMissingSlidePlaceholders: (slide: SlideData) => number;
/**
 * Adds the single placeholder of `type` that the layout defines and the slide
 * is missing — how an editor inserts a slide number, date or footer into a
 * slide without also restoring every other slot the author deleted.
 *
 * `type` is an ECMA-376 `ST_PlaceholderType` token (`sldNum`, `dt`, `ftr`,
 * `title`, `body`, …), matched as `findSlidePlaceholder` matches it. Returns
 * the slide's placeholder of that type — the one just added or the one already
 * there — or `null` when the layout reserves no such slot. Like the added slots
 * of `addMissingSlidePlaceholders`, it starts empty and inherits its geometry
 * and text style from the layout and master.
 * Pass `source: 'master'` to create an unmatched slot that inherits directly
 * from the master, without changing the layout (as the reference desktop app does when
 * demoting outline text into a Title Only slide). Returns null if absent there.
 */
declare const addSlidePlaceholder: (slide: SlideData, type: PlaceholderType, options?: {
  source?: "layout" | "master";
}) => SlideShapeData | null;
/**
 * Restore inherited text formatting on placeholders bound to the current
 * layout, including placeholders inside a group. Clears direct run, paragraph
 * and text-body appearance while preserving text, fields, hyperlinks,
 * language, outline levels and unknown extensions. Geometry and shape
 * appearance remain unchanged.
 * Returns the number of matching text placeholders processed.
 */
declare const resetSlidePlaceholderTextFormatting: (slide: SlideData) => number;
/**
 * Reset the current slide's layout: restore deleted slots, reset top-level
 * placeholder geometry and clear direct shape/text appearance to inherit the
 * layout again. Content, picture crops, relationships, IDs and unrelated shapes
 * remain intact.
 *
 * A placeholder inside a group has its appearance and text formatting reset
 * like any other, but keeps its geometry: where it sits is a statement about
 * the group's arrangement, and the layout has nothing to say about that.
 *
 * Returns the number of layout placeholders processed, including new slots.
 */
declare const resetSlideLayout: (slide: SlideData) => number;
//#endregion
//#region src/api/fn/shape-locks.d.ts
/** Whether the object's movement and resizing are both locked, as in the reference desktop app's Selection Pane. */
declare const isShapeLocked: (shape: SlideShapeData) => boolean;
/** Whether the reference desktop app preserves this object's aspect ratio while resizing it. */
declare const isShapeAspectRatioLocked: (shape: SlideShapeData) => boolean;
/**
 * Locks or unlocks object geometry like the reference desktop app's Selection Pane. Text,
 * selection, aspect-ratio constraints and extension data remain unchanged.
 * Group children keep their own locks; pass all descendants for Lock All.
 * Arrays are committed once per slide, including selections across slides.
 */
declare const setShapeLocked: (shapes: SlideShapeData | readonly SlideShapeData[], locked: boolean) => void;
/**
 * Sets the DrawingML `noChangeAspect` constraint without disturbing other
 * nonvisual properties. An absent constraint and an explicit false value both
 * read as unlocked; disabling the constraint removes only this attribute.
 */
declare const setShapeAspectRatioLocked: (shapes: SlideShapeData | readonly SlideShapeData[], locked: boolean) => void;
//#endregion
//#region src/api/fn/shape-custom-geometry.d.ts
/** One path to write; `fill` defaults to `'norm'` and `stroke` to `true`, as in the schema. */
interface CustomGeometryPathInput {
  /** Path coordinate-space width and height (`ST_PositiveCoordinate`). */
  readonly w: number;
  readonly h: number;
  readonly fill?: PathFillMode;
  readonly stroke?: boolean;
  readonly commands: readonly GeomCommand[];
}
interface CustomGeometryInput {
  readonly paths: readonly CustomGeometryPathInput[];
}
/**
 * Replaces the shape's geometry (preset or custom) with the given paths. The
 * text rectangle is the whole shape, as the reference desktop app writes for freeforms.
 * Coordinates are rounded to integers, which the schema requires.
 */
declare const setShapeCustomGeometry: (shape: SlideShapeData, geometry: CustomGeometryInput) => void;
//#endregion
//#region src/api/fn/shape-text-language.d.ts
/** Marks every run in the shape's text (and the end-of-paragraph marks) as `lang`, e.g. `en-US`. */
declare const setShapeTextLanguage: (shape: SlideShapeData, lang: string) => void;
/** The language of the shape's first run that names one, or `null`. */
declare const getShapeTextLanguage: (shape: SlideShapeData) => string | null;
/**
 * The literal `lang` of each run, field and line break in a paragraph, in
 * `getShapeParagraphElements` order, or `null` where the element names none.
 * Renderers use it for language-dependent layout such as the character a
 * decimal tab aligns on.
 */
declare const getParagraphElementLanguages: (target: SlideShapeData | TableCellData, paragraphIndex: number) => ReadonlyArray<string | null>;
//#endregion
//#region src/api/fn/shape-style.d.ts
/** One DrawingML style-list reference used by a shape quick style. */
interface ShapeStyleReference {
  /** 0 clears the reference; positive values index the theme format scheme. */
  readonly idx: number;
  /** Theme slot or sRGB color replacing the style list's `phClr`. */
  readonly color?: Color;
  /** Ordered DrawingML color transforms (for example the reference desktop app's shade). */
  readonly colorTransforms?: readonly ColorTransform[];
}
interface ShapeStyleFontReference {
  readonly idx: 'major' | 'minor' | 'none';
  readonly color?: Color;
  readonly colorTransforms?: readonly ColorTransform[];
}
/** Complete OOXML shape style. Direct text runs and geometry are untouched. */
interface ShapeStyleOptions {
  readonly line: ShapeStyleReference;
  readonly fill: ShapeStyleReference;
  readonly effect: ShapeStyleReference;
  readonly font: ShapeStyleFontReference;
}
/** Apply a complete native OOXML shape style while preserving geometry/text. */
declare const setShapeStyle: (shape: SlideShapeData, options: ShapeStyleOptions) => void;
//#endregion
//#region src/api/fn/shape-grouping.d.ts
/**
 * Groups two or more sibling shapes into a single `<p:grpSp>`,
 * returning the new group as a `SlideShapeData`. The group's
 * parent-space bounds are the union of its members' bounds; the members
 * keep their own relative position/size (nothing is rescaled). The
 * target slide is taken from the first shape — every shape must belong
 * to the same slide.
 *
 * Every shape must:
 *   - belong to the same slide as the others,
 *   - share the same immediate parent (the slide or a group),
 *   - have an explicit `<a:xfrm>` (placeholders that inherit position
 *     from the layout have none and can't be grouped),
 *   - appear at most once in `shapes` (grouping the same shape twice
 *     would duplicate its id).
 *
 * The group replaces its members at the position of the earliest one in
 * z-order, preserving member order regardless of selection order. Members
 * become one contiguous block, which can change interleaving with unselected shapes.
 */
declare const groupShapes: (shapes: ReadonlyArray<SlideShapeData>, opts?: {
  name?: string;
}) => SlideShapeData;
/**
 * Reverses `groupShapes`: removes the `<p:grpSp>` and re-inserts its
 * immediate children into its parent at the group's former position,
 * rescaling each child's own transform so it keeps its on-slide position
 * and size (matters when the group was moved/resized after creation, so
 * its `off`/`ext` diverged from its `chOff`/`chExt`). Returns the
 * children as fresh `SlideShapeData` handles, in their original order.
 * Group rotation and flips are composed with each child transform. Exact
 * preservation of rotated children requires uniform group scaling; anisotropic
 * scaling can introduce shear, which a child shape transform cannot represent.
 *
 * Throws if `group` isn't a group shape, or its `<p:grpSpPr>` carries no
 * `<a:xfrm>` (malformed — every authored group has one).
 */
declare const ungroupShapes: (group: SlideShapeData) => ReadonlyArray<SlideShapeData>;
//#endregion
//#region src/api/fn/media.d.ts
/** What `addSlideMedia` embeds or links. */
type SlideMediaSource = {
  kind: 'video';
  data: Uint8Array; /** Container of `data`. Detected from its signature when omitted. */
  format?: VideoFormat;
} | {
  kind: 'audio';
  data: Uint8Array; /** Container of `data`. Detected from its signature when omitted. */
  format?: AudioFormat;
} | {
  kind: 'online'; /** `http(s)` URL of the video. YouTube page URLs are rewritten to the embed form. */
  url: string;
};
type SlideMediaOptions = SlideMediaSource & {
  x: Emu;
  y: Emu;
  w: Emu;
  h: Emu;
  name?: string; /** Poster frame shown until the clip plays. A neutral play-button image when omitted. */
  poster?: Uint8Array; /** Format of `poster`. Detected from its signature when omitted. */
  posterFormat?: ImageFormat;
};
/** A picture's media, as `getShapeMedia` reports it. */
type ShapeMedia = {
  readonly kind: 'video' | 'audio';
  readonly partName: string;
  readonly contentType: string; /** Live view into the package part — copy before mutating. */
  readonly bytes: Uint8Array;
} | {
  /** Media linked by URL instead of embedded (an online video, or a linked file). */readonly kind: 'online';
  readonly url: string;
};
/**
 * Adds a video, an audio clip or an online (URL) video to the slide. Returns
 * the new picture shape — the clip's poster frame.
 *
 * Embedded clips are stored once per package: adding the same bytes again
 * reuses the existing `/ppt/media/mediaN.<ext>` part. The container format is
 * detected from the bytes (mp4 / m4v / mov / webm / avi / wmv, mp3 / wav / m4a /
 * ogg / wma); pass `format` for anything the signature check cannot tell.
 * Whether a clip plays depends on the codecs of the viewing application.
 *
 * The shape is an ordinary picture: `getShapeImageBytes` / `setShapeImage`
 * read and replace the poster, `getShapeMedia` reads the clip.
 */
declare const addSlideMedia: (slide: SlideData, opts: SlideMediaOptions) => SlideShapeData;
/**
 * Returns the clip behind a media picture, or `null` when the shape is not
 * one. Reads media authored by the reference desktop app, PptxGenJS and python-pptx alike: the
 * embedded part is taken from `<p14:media r:embed>` when present and from the
 * DrawingML `r:link` otherwise; a relationship with an external target is
 * reported as `'online'`.
 */
declare const getShapeMedia: (shape: SlideShapeData) => ShapeMedia | null;
/** Every shape on the slide that `getShapeMedia` reports a clip for. */
declare const findShapesWithMedia: (slide: SlideData) => ReadonlyArray<SlideShapeData>;
/** Durations removed from the source clip, in milliseconds. */
interface MediaTrim {
  readonly startMs: number;
  /** Duration removed from the end of the source clip. */
  readonly endMs: number;
}
interface MediaFade {
  readonly inMs: number;
  readonly outMs: number;
}
/** A named position in an embedded media clip, measured from its beginning. */
interface MediaBookmark {
  readonly name: string;
  readonly timeMs: number;
}
/**
 * How a clip plays in the slide show — the attributes of its
 * `<p:cMediaNode>` and the start condition of its time node.
 *
 * Trimming and fades are stored in the `p14:media` extension ([MS-PPTX]).
 * `trim.endMs` is the duration removed from the end of the clip, matching the
 * OOXML `p14:trim@end` meaning (it is not the playback end position).
 */
interface MediaPlayback {
  /** Starts with the slide instead of waiting for a click. */
  readonly autoplay: boolean;
  /** Milliseconds to wait after the slide starts before automatic playback. */
  readonly delayMs?: number;
  /** Plays again from the beginning until the slide moves on. */
  readonly loop: boolean;
  /** Playback volume, 0–1. The reference desktop app's own default is 0.8. */
  readonly volume: number;
  readonly muted: boolean;
  /** Video only: plays filling the screen. */
  readonly fullScreen: boolean;
  /** Hides the clip once it has played (`showWhenStopped="0"`). */
  readonly hideWhenStopped: boolean;
  /** Number of slides across which the clip should keep playing (`numSld`).
   * When absent, OOXML defaults to one slide. */
  readonly slideCount?: number;
  /** Returns playback to the beginning after natural completion (`fill="remove"`). */
  readonly rewindAfterPlaying?: boolean;
  readonly trim?: MediaTrim;
  readonly fade?: MediaFade;
  /** Named positions in the clip, in the order stored in the media extension. */
  readonly bookmarks?: readonly MediaBookmark[];
}
/**
 * Reads how the shape's clip plays, or `null` when the shape carries no media
 * time node — which is what a picture that is not a clip looks like, and also
 * what a clip pasted in without its node looks like (it shows no controls).
 */
declare const getShapeMediaPlayback: (shape: SlideShapeData) => MediaPlayback | null;
/**
 * Updates how the shape's clip plays. Omitted properties keep their current
 * value. Throws when the shape has no media time node — `addSlideMedia` writes
 * one, and a picture that is not a clip never plays.
 *
 * `delayMs` is a nonnegative safe integer and requires automatic playback.
 * Setting `autoplay: true` alone preserves an existing automatic delay;
 * switching from event playback starts immediately unless `delayMs` is supplied.
 *
 * `fullScreen` is a video attribute (`CT_TLMediaNodeVideo`); asking for it on
 * an audio clip throws rather than writing an attribute the schema rejects.
 */
declare const setShapeMediaPlayback: (shape: SlideShapeData, options: Partial<MediaPlayback>) => void;
//#endregion
//#region src/api/fn/shape-picture-style.d.ts
/**
 * Applies one of the reference desktop app's built-in Picture Styles (see
 * `BUILTIN_PICTURE_STYLES`) to a picture. Like the reference desktop app, it replaces the
 * picture's geometry, fill, border, effects and 3-D in `<p:spPr>` with the
 * style's own markup — which names no theme color, so the result does not
 * follow the theme — and keeps the picture, its crop and its position. Throws
 * when the shape is not a picture or `style` is not a built-in name.
 */
declare const setShapePictureStyle: (shape: SlideShapeData, style: BuiltinPictureStyleName) => void;
/**
 * The built-in picture style whose markup the picture's `<p:spPr>` carries
 * exactly (position, size and extensions aside), or `null`. A deck does not
 * record which style was applied, so a picture styled and then edited — a
 * different border, an extra effect — reads as `null`.
 */
declare const getShapePictureStyle: (shape: SlideShapeData) => BuiltinPictureStyleName | null;
/**
 * `ST_BlipCompression`: the target a picture was compressed for. The reference desktop app's
 * Compress Pictures writes `print` (220 ppi), `screen` (150 ppi) and `email`
 * (96 ppi); `hqprint` and `none` are the schema's other values.
 */
type ImageCompressionState = 'email' | 'screen' | 'print' | 'hqprint' | 'none';
/** The picture's `a:blip/@cstate`, or `null` when it has none. */
declare const getShapeImageCompressionState: (shape: SlideShapeData) => ImageCompressionState | null;
/**
 * Records what a picture was compressed for, as Compress Pictures does: it
 * sets `a:blip/@cstate` and adds the `a14:useLocalDpi` extension (without a
 * `val`, meaning true) that the reference desktop app writes alongside it. `null` removes
 * both. This only labels the picture; resample the pixels with
 * `setShapeImage`.
 */
declare const setShapeImageCompressionState: (shape: SlideShapeData, state: ImageCompressionState | null) => void;
//#endregion
//#region src/api/fn/shape-video.d.ts
/**
 * Restores a video to the reference desktop app's default shape formatting.
 *
 * This is deliberately restricted to video media pictures. It clears the
 * image recolor/correction elements and the shape-level fill, line, effects,
 * and 3-D style children, while preserving the poster/media relationship,
 * crop, transform, and unknown extension XML.
 */
declare const resetShapeVideoFormatting: (shape: SlideShapeData) => void;
//#endregion
//#region src/api/fn/slide-animation-edit.d.ts
/**
 * What to change about an effect. Every field is optional; the ones left out
 * keep what the slide already says.
 *
 * `build` is not an attribute flip: a paragraph build (`'allAtOnce'`,
 * `'byParagraph'`) replaces the effect with one per paragraph of the shape's
 * text, and `'asOneObject'` replaces the whole build with a single effect on
 * the shape; switching between the two paragraph builds replaces every
 * paragraph's effect. In every direction the id you passed stays on the effect
 * that takes the original's place in the click order, and the other effects of
 * a build are new — so a caller holding several handles into one build keeps
 * only the one it addressed.
 */
interface AnimationPatch {
  readonly effect?: AnimationEffect;
  /**
   * Which edge a fly, wipe or peek comes from or leaves by, or which corner
   * strips start from. Changing it rewrites the effect the same way changing
   * the preset does, because the direction is part of what the preset says.
   * Passing it for an effect that takes no direction is an error — including
   * one this patch is turning into such an effect. The same holds for the
   * four options below; `setShapeAnimation` documents which effects take
   * which.
   */
  readonly direction?: AnimationDirection;
  readonly orientation?: AnimationOrientation;
  readonly inOut?: AnimationInOut;
  readonly shape?: AnimationShape;
  readonly spokes?: number;
  /**
   * The emphasis effects' own options, as `setShapeAnimation` takes them:
   * Spin's direction and angle, Grow/Shrink's axes and size, Transparency's
   * amount, and the colour of the colour effects.
   */
  readonly spinDirection?: AnimationSpinDirection;
  readonly spinDegrees?: number;
  readonly scaleDirection?: AnimationScaleDirection;
  readonly scalePercent?: number;
  readonly transparencyPercent?: number;
  readonly color?: AnimationColor;
  /**
   * How long the effect runs. Left out, a changed effect keeps the duration
   * the slide gave it, unless the patch changes the preset itself: a new
   * preset takes its own default length, as picking one from the reference desktop app's
   * gallery does.
   */
  readonly durationMs?: number;
  readonly start?: AnimationStartCondition;
  readonly delayMs?: number;
  readonly build?: AnimationTextBuild;
}
/**
 * Changes one effect: its preset, its options, how long it runs, when it
 * starts, and how the shape's text is built.
 *
 * The effect keeps the id it was addressed by. Everything after it in the click
 * order is laid out again, because a longer effect moves what follows it and a
 * new start condition moves the effect itself between click stops.
 *
 * An effect whose duration the slide never stated gets this library's default
 * when it is rewritten; pass `durationMs` to say what it should be.
 */
declare const updateSlideAnimation: (slide: SlideData, id: number, patch: AnimationPatch) => void;
/**
 * Removes one effect. The effects after it move one place earlier in the click
 * order and are retimed to match.
 *
 * A paragraph of a build is removed on its own: the build entry the rest of the
 * paragraphs share stays, and goes only once the last of them is gone.
 */
declare const removeSlideAnimation: (slide: SlideData, id: number) => void;
/**
 * Moves one effect to `index` in the click order, counting only the slide's
 * main sequence — the order `getSlideAnimations` reports for steps whose
 * `sequence` is `'mainSeq'`.
 *
 * Each effect keeps the start condition it had, so moving a `'withPrevious'`
 * effect to the front does not turn it into a click: it becomes the first thing
 * that runs as the slide appears, which is what it already meant.
 */
declare const moveSlideAnimation: (slide: SlideData, id: number, index: number) => void;
//#endregion
//#region src/api/fn/_image-fill-layout.d.ts
type ImageTileAlignment = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
type ImageTileFlip = 'none' | 'x' | 'y' | 'xy';
/** Image placement inside its fill area; percentages use fractions (0.25 = 25%). */
type ImageFillLayout = {
  /** Omitted on write preserves the existing rotation setting. */rotateWithShape?: boolean;
} & ({
  mode: 'stretch';
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
} | {
  mode: 'tile';
  offsetX?: Emu;
  offsetY?: Emu;
  scaleX?: number;
  scaleY?: number;
  alignment?: ImageTileAlignment;
  flip?: ImageTileFlip;
});
//#endregion
//#region src/api/fn/slide-background.d.ts
/**
 * Reads back the slide's current background. Returns a discriminated
 * union mirroring `getShapeFill`'s shape, plus `inherit` when no
 * `<p:bg>` element is present (the slide picks up its background from
 * the layout / master).
 */
type SlideBackground = {
  readonly kind: 'solid';
  readonly color: string;
  readonly opacity?: number;
} | {
  readonly kind: 'gradient';
} | {
  readonly kind: 'pattern';
} | {
  readonly kind: 'image';
} | {
  readonly kind: 'inherit';
};
/**
 * Reads the slide's color-map override (`<p:clrMapOvr><p:overrideClrMapping/>`).
 * The mapping remaps the eight stable ECMA-376 color tokens (`bg1`,
 * `tx1`, `bg2`, `tx2`, `accent1`–`accent6`, `hlink`, `folHlink`) to
 * different theme positions. Returns `null` when the slide uses the
 * master's color map unchanged (the overwhelming common case).
 */
declare const getSlideColorMapOverride: (slide: SlideData) => Record<string, string> | null;
declare const getSlideBackground: (slide: SlideData) => SlideBackground;
/**
 * A simplified, render-ready view of one of the layout's non-placeholder
 * shapes. Resolves the bounds, preset, fill, and stroke without going
 * through the slide-bound `SlideShapeData` symbols. Returned by
 * `getSlideLayoutBackgroundShapes` for renderers that want to paint the
 * layout's brand-template decoration (corner bars, divider lines, logos
 * as solid rects, etc.) behind the slide's own shapes.
 */
interface SlideLayoutBackgroundShape {
  readonly kind: 'shape' | 'connector' | 'picture' | 'group' | 'graphicFrame';
  /** Bounds in EMU, or `null` when the shape inherits from its master. */
  readonly bounds: ShapeBounds | null;
  /** Preset geometry token for shapes (`'rect'`, `'roundRect'`, `'ellipse'`, …). */
  readonly preset: string | null;
  /** Fill color resolved to `#RRGGBB` (transforms + theme applied), or `null`. */
  readonly fillHex: string | null;
  /** Stroke color resolved to `#RRGGBB`, or `null`. */
  readonly strokeHex: string | null;
  /** Stroke width in EMU, or `null` when no explicit width is set. */
  readonly strokeWidthEmu: number | null;
  /** Rotation in degrees. */
  readonly rotation: number;
  /** Flip state. */
  readonly flip: {
    horizontal: boolean;
    vertical: boolean;
  };
}
/**
 * Returns the non-placeholder shapes on a layout as a render-ready
 * view. Useful for previewing brand-template decoration (corner bars,
 * background rectangles, divider lines) that would otherwise be hidden
 * because they aren't reachable through `getSlideLayoutPlaceholders`.
 *
 * Placeholders are excluded — they're better rendered through the
 * slide's own placeholder bounds (which already cascade through the
 * layout). Picture and group shapes are omitted; their bytes / nested
 * children would need the layout's relationship table to resolve.
 */
declare const getSlideLayoutBackgroundShapes: (pres: PresentationData, layout: SlideLayoutData) => ReadonlyArray<SlideLayoutBackgroundShape>;
/**
 * Non-placeholder decorative shapes on a slide layout — corner bars, divider
 * lines, **logos**, watermark text — as full `SlideShapeData` bound to the
 * layout part. Unlike `getSlideLayoutBackgroundShapes` (a flat summary that
 * drops pictures and groups), these work with every `getShape*` reader, so a
 * renderer can paint them with real geometry, fills, text and picture bytes.
 *
 * For reading / rendering only — the handles are bound to the layout part, not
 * a slide, so the mutating `setShape*` helpers should not be used on them.
 */
declare const getSlideLayoutShapes: (pres: PresentationData, layout: SlideLayoutData) => ReadonlyArray<SlideShapeData>;
/**
 * Non-placeholder decorative shapes on the layout's slide master — the
 * template decoration shared by every layout (background logos, divider lines,
 * etc.). Same render-ready `SlideShapeData` view as `getSlideLayoutShapes`.
 */
declare const getSlideMasterShapes: (pres: PresentationData, layout: SlideLayoutData) => ReadonlyArray<SlideShapeData>;
/**
 * Reads the slide layout's background. Same discriminated union as
 * `getSlideBackground` for slides — renderers fall back to this when
 * the slide's own background reports `'inherit'`. Walking one further
 * level to the master is left to callers (the same shape applies).
 */
/**
 * Reads the slide layout's pattern background when its `<p:bg>` is a
 * `<p:bgPr><a:pattFill>`. Same shape as `getSlideBackgroundPatternFill`
 * for slides.
 */
declare const getSlideLayoutBackgroundPatternFill: (pres: PresentationData, layout: SlideLayoutData, options?: {
  readonly preserveTheme?: boolean;
}) => {
  preset: string;
  foreground: string;
  background: string;
} | null;
/**
 * Reads the slide master's pattern background. Companion to
 * `getSlideLayoutBackgroundPatternFill`.
 */
declare const getSlideMasterBackgroundPatternFill: (pres: PresentationData, layout: SlideLayoutData, options?: {
  readonly preserveTheme?: boolean;
}) => {
  preset: string;
  foreground: string;
  background: string;
} | null;
/**
 * Reads the slide layout's gradient background from either `<p:bgPr><a:gradFill>`
 * or a theme style referenced by `<p:bgRef>`. Same shape as `getSlideBackgroundGradientFill`
 * for slides. Stop `resolvedColor` values include the presentation theme
 * and DrawingML color transforms.
 */
declare const getSlideLayoutBackgroundGradientFill: (layout: SlideLayoutData) => ReadGradientFill | null;
/**
 * Reads the slide master's gradient background when present. Mirrors
 * `getSlideLayoutBackgroundGradientFill`; returns `null` for any other
 * background kind. Useful for closing the bg cascade — slides that
 * report `'gradient'` inherit can now get the master's gradient
 * projected via `gradientDef`.
 * Stop `resolvedColor` values include the presentation theme and color transforms.
 */
declare const getSlideMasterBackgroundGradientFill: (pres: PresentationData, layout: SlideLayoutData) => ReadGradientFill | null;
/**
 * Reads the slide master's background. Same discriminated union as
 * `getSlideBackground` / `getSlideLayoutBackground`. Walks one rel up
 * from the layout to find the master, then reads `<p:bg>` (either
 * `<p:bgPr>` explicit fill or `<p:bgRef>` theme reference).
 *
 * Slides that report `'inherit'` on both their own background and the
 * layout's commonly inherit the brand fill from the master — this
 * closes the third level of the cascade.
 */
declare const getSlideMasterBackground: (pres: PresentationData, layout: SlideLayoutData) => SlideBackground;
declare const getSlideLayoutBackground: (layout: SlideLayoutData) => SlideBackground;
/**
 * Returns the gradient stops + path when the slide carries a
 * literal `<p:bgPr><a:gradFill>` or theme-referenced `<p:bgRef>` gradient.
 * Returns `null` for any other
 * background kind. Shape identical to `getShapeGradientFill` so renderers
 * can use the same projection logic for slide backgrounds.
 * Stop `resolvedColor` values include the presentation theme, slide color map
 * and DrawingML color transforms; `color` retains the authored token.
 */
declare const getSlideBackgroundGradientFill: (slide: SlideData) => ReadGradientFill | null;
/**
 * Returns the pattern preset + theme-resolved colors when the slide
 * carries a `<p:bgPr><a:pattFill>` background. Returns `null` for any
 * other background kind. Shape mirrors `getShapePatternFill`.
 */
declare const getSlideBackgroundPatternFill: (pres: PresentationData, slide: SlideData, options?: {
  readonly preserveTheme?: boolean;
}) => {
  preset: string;
  foreground: string;
  background: string;
} | null;
/**
 * Returns the embedded image bytes when the slide carries a
 * `<p:bgPr><a:blipFill>` background, or `null` for any other background
 * kind (solid / gradient / pattern / inherit) or when the `r:embed`
 * relationship points at an external `r:link` target.
 *
 * Companion to `getSlideBackground`, which only reports the
 * discriminated `kind`. Renderers that want to actually paint the
 * background image (preview generators, snapshot tools) call this.
 */
declare const getSlideBackgroundImageBytes: (slide: SlideData) => Uint8Array | null;
/**
 * Reads the slide layout's background image bytes when its `<p:bg>` is
 * a `<p:bgPr><a:blipFill>`. Same shape as `getSlideBackgroundImageBytes`
 * for slides. Returns `null` when the layout has no image background.
 */
declare const getSlideLayoutBackgroundImageBytes: (pres: PresentationData, layout: SlideLayoutData) => Uint8Array | null;
/**
 * Reads the slide master's background image bytes (via the layout's
 * master rel). Companion to `getSlideLayoutBackgroundImageBytes`.
 */
declare const getSlideMasterBackgroundImageBytes: (pres: PresentationData, layout: SlideLayoutData) => Uint8Array | null;
/** Sets a solid background; optional opacity is 0 (transparent) to 1 (opaque). */
declare const setSlideBackground: (slide: SlideData, color: Color, opacity?: number) => void;
/**
 * Replaces the slide background with a gradient. Uses the same stop and direction
 * settings as `setShapeGradientFill`; offsets and opacity range from 0 to 1.
 * Invalid settings leave the previous background unchanged.
 */
declare const setSlideBackgroundGradientFill: (slide: SlideData, options: GradientFillOptions) => void;
/** Natural background image size in EMUs, using fill DPI and embedded PNG/JPEG resolution.
 * Reads inherited backgrounds. Returns null for missing or unsupported image bytes.
 */
declare const getSlideBackgroundImageIntrinsicSize: (slide: SlideData) => {
  width: Emu;
  height: Emu;
} | null;
/** Reads direct or inherited background image crop fractions, including negative outsets.
 * Returns null when no source rectangle is specified or the background is not an image.
 */
declare const getSlideBackgroundImageCrop: (slide: SlideData) => ImageCrop | null;
/** Reads direct or inherited image background placement; returns null for other fills. */
declare const getSlideBackgroundImageFillLayout: (slide: SlideData) => ImageFillLayout | null;
/**
 * Changes background image placement while preserving media, crop and effects.
 * An inherited image becomes a slide override, leaving its layout/master unchanged.
 * Invalid settings or missing relationships leave the presentation unchanged.
 */
declare const setSlideBackgroundImageFillLayout: (slide: SlideData, layout: ImageFillLayout) => void;
/** Reads direct or inherited background picture opacity; null means no explicit opacity. */
declare const getSlideBackgroundImageOpacity: (slide: SlideData) => number | null;
/** Sets background image opacity (0–1), preserving media and placement. Null restores the default.
 * Inherited images become slide overrides. Invalid values leave the deck unchanged.
 */
declare const setSlideBackgroundImageOpacity: (slide: SlideData, opacity: number | null) => void;
/**
 * Updates a slide pattern background, preserving unspecified colors and transforms.
 * Inherited patterns become a slide override; the layout and master stay unchanged.
 * A new pattern uses the reference desktop app's defaults: pct5, accent1 foreground, bg1
 * background.
 */
declare const setSlideBackgroundPatternFill: (slide: SlideData, options: Partial<PatternFillOptions>) => void;
/**
 * Copies the source's effective background into a target slide, retaining its XML
 * and referenced parts. Inherited fills become explicit on the target. Theme
 * references use the target's theme; background graphics visibility is unchanged.
 * Across presentations, dependencies are copied with collision-free names.
 */
declare const copySlideBackground: (targetSlide: SlideData, sourceSlide: SlideData) => void;
/**
 * Applies the source slide's effective background throughout its presentation.
 * Like the reference desktop app on Mac, stores the fill on masters and clears slide/layout
 * overrides. Theme colors and image relationships are preserved without flattening the fill.
 */
declare const applySlideBackgroundToAll: (pres: PresentationData, slide: SlideData) => void;
/**
 * Sets a picture as the slide's background. Embeds `bytes` as a new
 * media part, wires a slide → image rel, and replaces any prior
 * background with a `<p:bgPr><a:blipFill><a:blip r:embed="…"/>
 * <a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr>` payload.
 *
 * Format is detected from magic bytes; pass `options.format` to
 * override.
 */
declare const setSlideBackgroundImage: (slide: SlideData, bytes: Uint8Array, options?: {
  format?: ImageFormat;
}) => void;
/** Clears any explicit slide background, restoring layout inheritance. */
declare const clearSlideBackground: (slide: SlideData) => void;
//#endregion
//#region src/api/fn/master-background-style.d.ts
/**
 * Choose one of the reference desktop app's twelve Background Styles (numbered row by row).
 * Changes the owning master, so every slide using that master inherits the
 * style. Existing slide and layout background overrides remain in place.
 * The theme must provide the referenced background fill style.
 */
declare const setSlideMasterBackgroundStyle: (slide: SlideData, style: number) => void;
/** A theme background preset; complex fill kinds without a gradient retain their kind. */
interface SlideMasterBackgroundStyle {
  readonly style: number;
  readonly selected: boolean;
  readonly fill: SlideBackground;
  readonly gradient: ReadGradientFill | null;
}
/**
 * Read the owning master's available Background Styles without changing the
 * presentation. Solid colors and gradient stop resolvedColor values include
 * the preset's color mapping and theme transforms. Missing theme rows are omitted.
 * Selection describes the master, independently of slide/layout overrides.
 */
declare const getSlideMasterBackgroundStyles: (slide: SlideData) => SlideMasterBackgroundStyle[];
//#endregion
//#region src/api/fn/color-map.d.ts
/**
 * The slide's effective color map: the master's `<p:clrMap>`, overlaid by a
 * layout or slide `<p:clrMapOvr><a:overrideClrMapping>` when present. An explicit
 * slide `masterClrMapping` bypasses the layout override. Falls back to
 * the standard map for decks that omit it. Given a layout, it is the map the
 * layout itself is drawn with (its master's, under its own override).
 *
 * Pass the result to color resolution / renderers so `schemeClr` tokens map to
 * the theme slot the reference desktop app actually paints — critical for decks with an
 * inverted map (`bg1="dk1" tx1="lt1"`).
 */
declare const getEffectiveColorMap: (slide: SlideData | SlideLayoutData) => Record<string, string>;
/**
 * The concrete `#RRGGBB` color the deck paints body text in, resolved through
 * the effective color map + theme. Read from the master's `bodyStyle` level-1
 * default run properties; falls back to the `tx1` token resolved through the
 * map. Returns `null` only when the deck carries no theme.
 *
 * Tables and charts authored by this library inherit no master text style, so
 * their text otherwise falls back to `tx1` — which an inverted color map paints
 * the SAME as the background, making the text invisible. Baking the body color
 * in keeps generated tables / charts readable on whatever surface the deck uses.
 */
declare const resolveDeckBodyTextColor: (slide: SlideData | SlideLayoutData) => string | null;
//#endregion
//#region src/api/fn/slide-transition.d.ts
/**
 * Reads back the slide's current transition (or `null` if no
 * `<p:transition>` is present). The returned shape mirrors what
 * `setSlideTransition` accepts.
 */
declare const getSlideTransition: (slide: SlideData) => SlideTransition | null;
/**
 * Sets the slide's transition effect. A sound set with
 * `setSlideTransitionSound` is kept, as the reference desktop app keeps it when the effect
 * changes.
 */
declare const setSlideTransition: (slide: SlideData, options: TransitionOptions) => void;
/** Removes any existing transition on the slide, including its sound. */
declare const clearSlideTransition: (slide: SlideData) => void;
/** What plays when the slide's transition starts. */
type TransitionSound = {
  readonly kind: 'stop';
} | {
  readonly kind: 'play';
  readonly name: string;
  readonly loop: boolean;
};
/**
 * `stop` ends any sound still playing from an earlier slide ("Stop Previous
 * Sound"); `play` embeds a WAV file, the only format `p:snd` allows
 * (ECMA-376 §19.5.66, CT_EmbeddedWAVAudioFile).
 */
type TransitionSoundInput = {
  readonly kind: 'stop';
} | {
  readonly kind: 'play';
  readonly data: Uint8Array;
  readonly name: string;
  readonly loop?: boolean;
};
/** Reads the transition's sound, or `null` when it has none. */
declare const getSlideTransitionSound: (slide: SlideData) => TransitionSound | null;
/**
 * Sets (or with `null` removes) the sound played when the slide's transition
 * starts. A slide without a transition gets one with no effect, which is how
 * the reference desktop app stores a sound on a "None" transition.
 */
declare const setSlideTransitionSound: (slide: SlideData, sound: TransitionSoundInput | null) => void;
//#endregion
//#region src/api/fn/slide-notes.d.ts
interface SlideNotesTextFormatOptions {
  /** Resolve scheme colors through the notes master theme and color map. */
  readonly resolveColors?: boolean;
}
/** Resolve a notes text color token through the notes master theme and map.
 * Literal tokens remain unchanged in the notes format readers; this helper is
 * for transient editor display state (such as a pending typing format).
 */
declare const resolveSlideNotesTextColor: (slide: SlideData, color: string) => string | null;
/**
 * Returns the slide's speaker notes (`null` if none). Pulls plain text
 * from the `body` placeholder; multi-line notes use `\n`.
 */
declare const getSlideNotes: (slide: SlideData) => string | null;
/** Returns the kind and UTF-16 position of each visible notes line separator. */
declare const getSlideNotesLineBreaks: (slide: SlideData) => ReadonlyArray<{
  readonly position: number;
  readonly kind: "paragraph" | "break";
}>;
/** A direct character-format range in a slide's speaker notes. */
interface SlideNotesTextFormatRange {
  readonly start: number;
  readonly end: number;
  readonly format: ReadTextFormat;
}
/**
 * Reads run formatting from the notes body using the same UTF-16 offsets as
 * `setSlideNotesFormat`. By default scheme colors remain literal references;
 * pass `resolveColors` to resolve them through the notes master theme and
 * color map. Paragraph and line-break separators are included in the offsets
 * but are not returned as formatted ranges.
 */
declare const getSlideNotesTextFormats: (slide: SlideData, options?: SlideNotesTextFormatOptions) => ReadonlyArray<SlideNotesTextFormatRange>;
/**
 * Reads the paragraph-end character format used for typing at a notes caret.
 * By default scheme colors remain literal references; pass `resolveColors` to
 * resolve them through the notes master theme and color map.
 */
declare const getSlideNotesParagraphEndFormat: (slide: SlideData, paragraphIndex: number, options?: SlideNotesTextFormatOptions) => ReadTextFormat;
/**
 * Sets the slide's speaker notes. Creates the `notesSlide` part and
 * wires up the rels (slide ↔ notesSlide ↔ notesMaster) on first call;
 * subsequent calls just replace the body placeholder text.
 */
/**
 * Returns every slide whose `<p:sldId show="0">` flag is set —
 * complement of `getVisibleSlides`. Useful for audit UIs and
 * batch-unhide operations.
 */
declare const getHiddenSlides: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide in document order whose `<p:sldId show="0">`
 * flag is *not* set. Convenience over `getSlides(pres).filter(s =>
 * !isSlideHidden(s))` — useful when an export pipeline needs to
 * skip hidden slides without touching the `show` attribute itself.
 */
declare const getVisibleSlides: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide carrying at least one chart graphic frame.
 * Built on `isChartShape`.
 */
declare const getSlidesWithCharts: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Filters `getSlidesWithCharts` to slides carrying at least one chart
 * of the given `kind` (`'bar'`, `'column'`, `'line'`, `'pie'`,
 * `'doughnut'`, `'area'`). Reads each chart's spec via `getSlideCharts`
 * so the predicate respects what the renderers actually see.
 */
declare const findSlidesWithChartKind: (pres: PresentationData, kind: ChartKind) => ReadonlyArray<SlideData>;
/**
 * Presentation-level version of `findChartsWithTrendlines`. Returns
 * every slide carrying at least one chart with a trendline on any
 * series. Useful for "audit every trendline in this deck" workflows
 * before publishing.
 */
declare const findSlidesWithChartTrendlines: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide where at least two shapes have overlapping
 * bounding boxes. Built on `findOverlappingShapePairs`. Useful for
 * deck-wide layout audits — surfacing slides that may have stacked
 * or accidentally-colliding content for human review.
 */
declare const getSlidesWithOverlap: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide carrying at least one table graphic frame.
 * Built on `isTableShape`.
 */
declare const getSlidesWithTables: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide carrying at least one image-bearing shape
 * (a `<p:pic>` picture or a regular shape with `<a:blipFill>`).
 * Built on `hasShapeImage`.
 */
declare const getSlidesWithImages: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide that has at least one comment attached.
 * Convenience over `getSlides(pres).filter(s =>
 * getSlideComments(s).length > 0)`.
 */
declare const getSlidesWithComments: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * One entry per slide with non-empty notes, carrying its 0-based
 * slide index and the notes text. Useful for "export speaker
 * notes to a separate document" workflows that need both the
 * notes and the slide they belong to.
 */
interface PresentationNotesEntry {
  readonly slideIndex: number;
  readonly notes: string;
}
/**
 * One entry per chart in the deck, carrying both the chart and the
 * 0-based slide it was attached to.
 */
interface PresentationChartEntry {
  readonly slideIndex: number;
  readonly chart: SlideChartData;
}
/**
 * One entry per image-bearing shape in the deck, carrying the
 * shape (picture or image-filled) and the 0-based slide it lives
 * on. Sibling of `getAllCharts` / `getAllTables`.
 */
interface PresentationImageEntry {
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
}
/**
 * Returns every image-bearing shape across the deck (pictures and
 * shapes with `<a:blipFill>`), paired with its 0-based slide
 * index. Built on `hasShapeImage`.
 */
declare const getAllImages: (pres: PresentationData) => ReadonlyArray<PresentationImageEntry>;
/**
 * One entry per table in the deck, carrying the table shape and
 * the 0-based slide it sits on. Sibling of `getAllCharts`.
 */
interface PresentationTableEntry {
  readonly slideIndex: number;
  readonly table: SlideShapeData;
}
/**
 * Returns every table across every slide in the deck, paired with
 * the 0-based index of its slide. Built on `isTableShape`.
 */
declare const getAllTables: (pres: PresentationData) => ReadonlyArray<PresentationTableEntry>;
/**
 * Returns every chart across every slide in the deck, paired with
 * the 0-based index of its slide. Useful for chart-inventory UIs
 * and bulk chart-update pipelines.
 */
declare const getAllCharts: (pres: PresentationData) => ReadonlyArray<PresentationChartEntry>;
/**
 * One entry per external hyperlink found in a shape's text body,
 * carrying the URL, the linked shape, and the 0-based slide index.
 * Each hyperlinked shape is reported once (the URL of its first
 * `<a:hlinkClick>` run).
 */
interface PresentationHyperlinkEntry {
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
  readonly url: string;
}
/**
 * Returns every external hyperlink in the deck — one entry per
 * shape whose text body carries an `<a:hlinkClick>`. Useful for
 * "link audit" passes before publishing, and for building a
 * deck-wide table of contents of outbound URLs.
 */
declare const getAllHyperlinks: (pres: PresentationData) => ReadonlyArray<PresentationHyperlinkEntry>;
/**
 * Dense per-slide hyperlink count array. Counts shapes whose
 * `getShapeHyperlink` is non-null, mirroring `getAllHyperlinks` but
 * cheaper when the caller only wants per-slide counts. Rounds out the
 * deck-density family alongside the chart / table / image / shape /
 * text / comment counters.
 */
declare const getPresentationHyperlinkCountsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * Returns every distinct external URL referenced by any shape in
 * the deck, in first-seen order. Sibling of `getAllHyperlinks`
 * (which keeps duplicates and slide indices). Useful for "are
 * these URLs all live?" audits where checking each URL once is
 * enough.
 */
declare const getDistinctHyperlinkUrls: (pres: PresentationData) => ReadonlyArray<string>;
/**
 * Returns every slide carrying at least one shape with an external
 * hyperlink. Built on `findHyperlinkedShapes`. Useful for navigation
 * UIs that want to surface only the slides containing outbound URLs.
 */
declare const getSlidesWithHyperlinks: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Returns every slide containing at least one shape whose external
 * hyperlink matches `needle` (substring or `RegExp`). Sibling of
 * `findSlidesByText` for outbound-URL audits — e.g. "every slide
 * that links to old.docs.example.com".
 */
declare const findSlidesByHyperlink: (pres: PresentationData, needle: string | RegExp) => ReadonlyArray<SlideData>;
/**
 * Bulk URL migration. Re-points every shape across the deck whose
 * first hyperlink exactly equals `from` to instead point at `to`.
 * Returns the number of shapes updated. Built on
 * `setShapeHyperlink`, so each update goes through the standard
 * rels-allocation path and stays schema-valid.
 *
 * Matching is exact (case-sensitive). To migrate by pattern, use
 * `findSlidesByHyperlink` to locate slides and rewrite each shape
 * yourself.
 */
declare const replaceHyperlink: (pres: PresentationData, from: string, to: string) => number;
/**
 * Removes every external hyperlink across the deck — useful for
 * sanitizing a template before sharing, or for stripping outbound
 * URLs from an exported preview. Returns the number of shapes
 * cleared. Each call goes through `setShapeHyperlink(_, null)`.
 */
declare const clearAllHyperlinks: (pres: PresentationData) => number;
/**
 * Slide-scoped sibling of `clearAllHyperlinks`. Removes every
 * external hyperlink on this slide and returns the number of
 * shapes cleared.
 */
declare const clearSlideHyperlinks: (slide: SlideData) => number;
/**
 * One entry per comment in the deck, carrying both the comment and
 * the 0-based slide it was attached to.
 */
interface PresentationCommentEntry {
  readonly slideIndex: number;
  readonly comment: SlideCommentData;
}
/**
 * Returns every comment across every slide in the deck, each paired
 * with the 0-based index of its slide. Useful for review-summary
 * UIs that show all annotations in one chronological list.
 */
declare const getAllComments: (pres: PresentationData) => ReadonlyArray<PresentationCommentEntry>;
/**
 * Returns every slide's speaker notes alongside its 0-based index.
 * Skips slides whose notes are empty / unset.
 */
declare const getAllNotes: (pres: PresentationData) => ReadonlyArray<PresentationNotesEntry>;
/**
 * Returns every slide in the presentation that carries non-empty
 * speaker notes. Convenience over `getSlides(pres).filter(s =>
 * getSlideNotes(s) !== null && getSlideNotes(s) !== '')`.
 */
declare const getSlidesWithNotes: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Predicate sibling of `getSlideNotes`. Returns `true` when the
 * slide carries a non-empty `notesSlide` body — i.e. whatever
 * `getSlideNotes(slide)` would surface is a non-empty string.
 *
 * Cheap to call in hot loops where the caller only needs to know
 * "are there any notes here?" without materializing the text.
 */
declare const hasSlideNotes: (slide: SlideData) => boolean;
/**
 * Code-point length of the slide's speaker notes, or `0` when the
 * slide has no notes. Counts via `Array.from`, so surrogate-pair
 * characters (emoji, supplementary CJK) count as 1 — matches
 * `getSlideTextLength`.
 */
declare const getSlideNotesLength: (slide: SlideData) => number;
/**
 * Concatenated speaker notes from every slide, joined with the
 * given `separator` (defaults to a form-feed, `\f`). Slides with
 * no notes contribute the empty string. Sibling of
 * `getPresentationText` — useful for search-indexing notes
 * across a whole deck.
 */
declare const getPresentationNotesText: (pres: PresentationData, separator?: string) => string;
/**
 * Total code-point length of speaker notes across every slide.
 * Sibling of `getPresentationTextLength`; counts surrogate-pair
 * characters as 1 each. Cheaper than `getPresentationNotesText`
 * when the caller only needs the size.
 */
declare const getPresentationNotesLength: (pres: PresentationData) => number;
/**
 * Dense per-slide notes-length array. Code-point length per
 * `getSlideNotesLength` (`0` when the slide has no notes part). Pair
 * with `getPresentationTextLengthsBySlide` for handout / talk-track
 * audits — slides with no on-screen text but heavy speaker notes are
 * often the slow part of a presentation.
 */
declare const getPresentationNotesLengthsBySlide: (pres: PresentationData) => ReadonlyArray<number>;
/**
 * Returns every slide whose `<p:spTree>` carries no shapes (per
 * `getSlideShapes`). Useful as a "find the section dividers I forgot
 * to fill in" pre-publish check.
 */
declare const getEmptySlides: (pres: PresentationData) => ReadonlyArray<SlideData>;
/**
 * Appends `text` to the slide's existing notes on its own line.
 * Equivalent to `setSlideNotes(slide, (getSlideNotes(slide) ?? '') + '\n' + text)`,
 * minus the leading newline when there were no notes yet.
 */
declare const appendSlideNotes: (slide: SlideData, text: string) => void;
/**
 * Replaces speaker notes. With `range`, replaces that UTF-16 selection while
 * retaining untouched run/paragraph XML and inheriting the insertion format.
 * When inserting line-feed characters through a ranged edit, `newlines`
 * selects whether those inserted separators become paragraph breaks
 * (`'paragraph'`, the default) or paragraph-internal soft breaks (`'break'`).
 * Existing separators outside the edited range keep their original OOXML
 * kind; the option does not normalize the rest of the notes.
 * `preserveFormatting` infers one changed range from the common prefix/suffix;
 * apply disjoint edits separately to retain the formatting between them.
 */
declare const setSlideNotes: (slide: SlideData, value: string, options?: {
  preserveFormatting?: boolean;
  range?: {
    start: number;
    end: number;
  };
  newlines?: "paragraph" | "break";
}) => void;
/** Applies Change Case, as the reference desktop app does, to a speaker-notes range. */
declare const transformSlideNotesCase: (slide: SlideData, value: TextCase, options?: {
  range?: {
    start: number;
    end: number;
  };
}) => void;
/**
 * Applies direct run formatting to a UTF-16 range in the slide's speaker notes.
 * Unselected runs and paragraph XML are preserved. An empty range is a no-op.
 * Use `paragraphEnd` instead of `range` to update the paragraph-end character
 * formatting for subsequent typing, including in an empty notes body.
 */
declare const setSlideNotesFormat: (slide: SlideData, format: TextFormat, options: {
  range: {
    start: number;
    end: number;
  };
  reset?: boolean;
  paragraphEnd?: never;
} | {
  paragraphEnd: number;
  reset?: boolean;
  range?: never;
}) => void;
/**
 * Removes the slide's speaker-notes part entirely. Drops the
 * `notesSlide` part + its `.rels`, and unwires the slide → notesSlide
 * relationship. No-op when the slide has no notes.
 *
 * The shared `notesMaster` part is left alone; other slides may still
 * reference it.
 */
declare const removeSlideNotes: (slide: SlideData) => void;
/**
 * Removes the speaker-notes part from every slide that has one.
 * Built on `removeSlideNotes`. Returns the number of slides
 * stripped. Useful as a privacy/sharing helper before exporting a
 * deck whose notes contain internal commentary.
 */
declare const clearAllSlideNotes: (pres: PresentationData) => number;
//#endregion
//#region src/api/fn/slide-size.d.ts
/**
 * Width × height of the slide canvas, in EMU. `type` is the reference desktop app's
 * aspect-ratio hint (`screen4x3`, `screen16x9`, ...); the actual size
 * is always `width` × `height`.
 */
interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
  readonly type?: string;
}
/** Returns the slide canvas size, or `null` if `presentation.xml` omits it. */
declare const getSlideSize: (pres: PresentationData) => SlideSize | null;
/**
 * The number the reference desktop app prints on the first slide — `<p:presentation
 * firstSlideNum="N"/>`, or `1` when the deck does not say. A `slidenum` field
 * on the n-th slide shows this plus n − 1, so a renderer that substitutes live
 * field values has to start counting here rather than at one.
 */
declare const getPresentationFirstSlideNumber: (pres: PresentationData) => number;
/**
 * Sets the slide canvas size. Creates `<p:sldSz>` when absent, replaces
 * its attributes when present. Dimensions must be 1–56 inches in EMU;
 * invalid dimensions throw before changing the presentation.
 * The `type` hint is preserved as given. `content: 'fit'` uniformly scales
 * content to fit the new page and centers it, including layout/master artwork
 * and physical text formatting. The default keeps existing content unchanged.
 *
 * Schema ordering: `<p:sldSz>` follows `<p:sldIdLst>` per ECMA-376
 * §19.2.1.26; we insert at the correct position when bootstrapping.
 */
declare const setSlideSize: (pres: PresentationData, opts: SlideSize, options?: {
  content?: "keep" | "fit";
}) => void;
/** 10in × 7.5in (`screen4x3`). */
declare const SLIDE_SIZE_4_3: SlideSize;
/** 13.333in × 7.5in (`screen16x9`) — the default since 2013. */
declare const SLIDE_SIZE_16_9: SlideSize;
/** 13.333in × 8.33in (`screen16x10`). */
declare const SLIDE_SIZE_16_10: SlideSize;
//#endregion
//#region src/api/fn/guides.d.ts
/** A presentation-level drawing guide. Position is in EMU from the left/top edge. */
interface DrawingGuide {
  id: number;
  axis: 'x' | 'y';
  position: number;
  color: string;
}
/** Returns null when the presentation has no explicit guide list; [] means cleared. */
declare function getDrawingGuides(presentation: PresentationData): DrawingGuide[] | null;
/** Replaces presentation guides, retaining metadata/extensions on surviving guide IDs. */
declare const setDrawingGuides: (presentation: PresentationData, guides: readonly DrawingGuide[]) => void;
/** Whether drawing guides are visible in the presentation's slide editing view. */
declare function getDrawingGuidesVisible(presentation: PresentationData): boolean;
/** Persists slide-view guide visibility without changing other view preferences. */
declare const setDrawingGuidesVisible: (presentation: PresentationData, visible: boolean) => void;
/** Document grid snapping; null means no slide-view preferences are stored. */
declare function getSnapToGrid(presentation: PresentationData): boolean | null;
/** Turns snapping objects to the document grid (View ▸ Guides ▸ Snap to Grid) on or off. */
declare const setSnapToGrid: (presentation: PresentationData, enabled: boolean) => void;
/** Grid spacing in EMU, or null when no document spacing is specified. */
declare function getGridSpacing(presentation: PresentationData): {
  x: number;
  y: number;
} | null;
/** Sets the document grid's horizontal (`x`) and vertical (`y`) spacing, in EMU. */
declare const setGridSpacing: (presentation: PresentationData, spacing: {
  x: number;
  y: number;
}) => void;
/** Slides whose bodies are collapsed in the reference desktop app's outline view, in deck order. */
declare function getCollapsedOutlineSlides(presentation: PresentationData): readonly SlideData[];
/** Persists outline collapse and retains other view settings. */
declare const setSlideOutlineCollapsed: (slide: SlideData | readonly SlideData[], collapsed: boolean) => void;
//#endregion
//#region src/api/fn/shape-image-fill-layout.d.ts
/** Reads direct picture/image-fill placement; returns null when there is no image fill. */
declare const getShapeImageFillLayout: (shape: SlideShapeData) => ImageFillLayout | null;
/** Changes placement without replacing media, crop, or effects. Invalid inputs leave the shape unchanged. */
declare const setShapeImageFillLayout: (shape: SlideShapeData, layout: ImageFillLayout) => void;
/**
 * Natural image-fill size before tile scaling, in EMUs. Reads PNG/JPEG dimensions,
 * the fill's DPI override, and PNG pHYs/JPEG JFIF resolution. Missing resolution
 * uses 96 DPI. Returns null for unavailable bytes or unsupported image headers.
 */
declare const getShapeImageIntrinsicSize: (shape: SlideShapeData) => {
  width: Emu;
  height: Emu;
} | null;
//#endregion
//#region src/api/fn/show-properties.d.ts
interface SlideShowProperties {
  mode: {
    kind: 'present';
  } | {
    kind: 'browse';
    showScrollbar: boolean;
  } | {
    kind: 'kiosk';
    restart: number;
  };
  slides: {
    kind: 'all';
  } | {
    kind: 'range';
    start: number;
    end: number;
  } | {
    kind: 'customShow';
    id: number;
  };
  loop: boolean;
  showNarration: boolean;
  showAnimation: boolean;
  useTimings: boolean;
  /**
   * Show Media Controls: whether hovering a video or audio clip in the show
   * reveals its play bar. The reference desktop app stores it as `p14:showMediaCtrls`
   * in the show properties' extension list; omit it to leave the file as is.
   */
  showMediaControls?: boolean;
}
/** Reads presentation-wide slideshow settings using the OOXML defaults. */
declare function getSlideShowProperties(pres: PresentationData): SlideShowProperties;
/** Replaces slideshow settings while preserving print/web options, pen color, and extensions. */
declare function setSlideShowProperties(pres: PresentationData, settings: SlideShowProperties): void;
//#endregion
//#region src/api/index.d.ts
declare const VERSION: string;
//#endregion
export { getSlideNotes as $, findShapesWithAnimation as $a, ChartUpDownBars as $c, getSlideLayout as $d, loadPresentation as $f, SlideLayoutPlaceholder as $i, getShapeTextAutoFitParams as $l, resetSlidePlaceholderTextFormatting as $n, slidesUsingMediaPart as $o, BulletStyle as $p, searchSlides as $r, setTableStyleId as $s, getSlideMasterBackgroundImageBytes as $t, setShapeRotation as $u, PresentationNotesEntry as A, getShapeParagraphElements as Aa, RadarChartSpec as Ac, getShapeStrokeEffective as Ad, getShapeSize as Af, getSlideMasterName as Ai, getSlideComments as Al, asColor as Am, getShapeMedia as An, getShapeImagePartName as Ao, MasterPlaceholderType as Ap, getPresentationTextLength as Ar, getTableSize as As, setSlideMasterBackgroundStyle as At, getShapeEffectsEffective as Au, getAllCharts as B, setParagraphLevel as Ba, ChartDateAxis as Bc, findShapesByHyperlink as Bd, pointInShape as Bf, setSlideMasterPreserved as Bi, setShapeHyperlink as Bl, inches as Bm, getShapeTextLanguage as Bn, MediaPart as Bo, GeomPoint as Bp, getSlideText as Br, setTableCellAlignment as Bs, getSlideBackgroundImageFillLayout as Bt, getShapeStrokeArrow as Bu, getPresentationFirstSlideNumber as C, getParagraphEndFormat as Ca, ComboChartSpec as Cc, getShapeFillEffective as Cd, getShapeKind as Cf, setHandoutSlidesPerPage as Ci, getCommentStatus as Cl, GlowOptions as Cm, MediaPlayback as Cn, getShapeImageContrast as Co, MorphOption as Cp, findSlideByTitle as Cr, getTableCellText as Cs, getSlideTransitionSound as Ct, ShapeEffect as Cu, PresentationCommentEntry as D, getParagraphSpacing as Da, Pie3DChartSpec as Dc, getShapeStrokeColor as Dd, getShapePosition as Df, addSlideLayoutPlaceholder as Di, getPresentationCommentCountsBySlide as Dl, Color as Dm, SlideMediaSource as Dn, getShapeImageFormat as Do, TransitionPreset as Dp, getOutlineText as Dr, getTableColumnWidths as Ds, resolveDeckBodyTextColor as Dt, getShape3D as Du, PresentationChartEntry as E, getParagraphLineSpacing as Ea, OfPieChartSpec as Ec, getShapeStrokeCap as Ed, getShapePlaceholderType as Ef, addSlideLayout as Ei, getPresentationCommentCountsByAuthor as El, ShadowOptions as Em, SlideMediaOptions as En, getShapeImageFillBytes as Eo, TransitionOptions as Ep, getAllShapes as Er, getTableCells as Es, getEffectiveColorMap as Et, findShapesByEffect as Eu, clearAllSlideNotes as F, getShapeRunText as Fa, isChartSpec as Fc, findEmptyPlaceholders as Fd, getSlideMasterUsageCounts as Ff, setSlideLayoutBackgroundGraphicsHidden as Fi, ParagraphProperties as Fl, VideoFormat as Fm, ShapeStyleFontReference as Fn, setShapeImageBrightness as Fo, BuiltinPictureStyleName as Fp, getSlideIndex as Fr, isChartShape as Fs, copySlideBackground as Ft, setShapeShadow as Fu, getAllTables as G, setShapeRunHyperlink as Ga, ChartManualLayout as Gc, findShapesWithHyperlinks as Gd, setShapeHidden as Gf, getSlideLayoutName as Gi, TextAnchor as Gl, isShapeAspectRatioLocked as Gn, getMediaParts as Go, LightRig as Gp, isSlideBackgroundGraphicsHidden as Gr, setTableCellMargins as Gs, getSlideLayoutBackground as Gt, setShapeFlip as Gu, getAllHyperlinks as H, setParagraphSpacing as Ha, ChartErrorBars as Hc, findShapesByName as Hd, setShapeAltTitle as Hf, findSlideLayout as Hi, getShapeRunFormat as Hl, pt as Hm, CustomGeometryInput as Hn, PresentationSummary as Ho, Bevel as Hp, getSlideXmlString as Hr, setTableCellBorders as Hs, getSlideBackgroundImageOpacity as Ht, getShapeStrokeSketch as Hu, clearSlideHyperlinks as I, isParagraphBulletPicture as Ia, ChartAxisScaling as Ic, findShapeById as Id, getUnusedSlideMasters as If, setSlideLayoutFootersIncluded as Ii, ParagraphTabStop as Il, ImageFormat as Im, ShapeStyleOptions as In, setShapeImageContrast as Io, BUILTIN_PICTURE_STYLES as Ip, getSlideInfo as Ir, isTableShape as Is, getSlideBackground as It, setShapeSoftEdge as Iu, getHiddenSlides as J, getSlideTitle as Ja, ChartSeries as Jc, findSlidePlaceholders as Jd, getSlideSections as Jf, getSlideLayoutType as Ji, appendShapeText as Jl, setShapeLocked as Jn, getPresentationSummary as Jo, PresetMaterial as Jp, replaceTextInNotes as Jr, setTableCellTextDirection as Js, getSlideLayoutBackgroundPatternFill as Jt, setShapeNoFill as Ju, getDistinctHyperlinkUrls as K, setShapeRunText as Ka, ChartOfPie as Kc, findSlidePlaceholder as Kd, shapesOverlap as Kf, getSlideLayoutPartName as Ki, TextAutoFit as Kl, isShapeLocked as Kn, getOrphanMediaPartNames as Ko, LightRigDirection as Kp, isSlideHidden as Kr, setTableCellParagraphs as Ks, getSlideLayoutBackgroundGradientFill as Kt, setShapeGradientFill as Ku, findSlidesByHyperlink as L, setParagraphAlignment as La, ChartDataLabelPosition as Lc, findShapeByName as Ld, isShapeHidden as Lf, setSlideLayoutTitleIncluded as Li, getParagraphPropertiesEffective as Ll, Emu as Lm, ShapeStyleReference as Ln, setShapeImageCrop as Lo, CustomGeometry as Lp, getSlideLayoutCount as Lr, mergeTableCells as Ls, getSlideBackgroundGradientFill as Lt, clearShapeFill as Lu, SlideNotesTextFormatOptions as M, getShapeRunCount as Ma, ScatterChartSpec as Mc, getShapeStrokeJoin as Md, getSlideMasterCount as Mf, isSlideMasterPreserved as Mi, setCommentStatus as Ml, PresetShape as Mm, setShapeMediaPlayback as Mn, hasShapeText as Mo, BuiltinTableStyle as Mp, getShapeAt as Mr, getTableStyleId as Ms, SlideLayoutBackgroundShape as Mt, setShapeGlow as Mu, appendSlideNotes as N, getShapeRunHyperlink as Na, StockChartSpec as Nc, getShapeStrokeOpacity as Nd, getSlideMasterPartName as Nf, removeSlideLayout as Ni, setCommentText as Nl, PlaceholderType as Nm, groupShapes as Nn, isShapeImageGrayscale as No, BuiltinTableStyleName as Np, getSlideAt as Nr, insertTableColumn as Ns, applySlideBackgroundToAll as Nt, setShapeInnerShadow as Nu, PresentationHyperlinkEntry as O, getShapeHyperlinkTooltip as Oa, PieChartSpec as Oc, getShapeStrokeColorResolved as Od, getShapePreset as Of, addSlideMaster as Oi, getPresentationCommenters as Ol, HexColor as Om, addSlideMedia as On, getShapeImageLinkUrl as Oo, HandoutMasterPlaceholderType as Op, getPresentationShapeCountsBySlide as Or, getTableDimensions as Os, SlideMasterBackgroundStyle as Ot, getShapeEffect as Ou, clearAllHyperlinks as P, getShapeRunHyperlinkTooltip as Pa, SurfaceChartSpec as Pc, getShapeStrokeWidth as Pd, getSlideMasterPartNames as Pf, removeSlideMaster as Pi, ParagraphBulletDetail as Pl, AudioFormat as Pm, ungroupShapes as Pn, resetShapeImageColorEffects as Po, BUILTIN_TABLE_STYLES as Pp, getSlideCount as Pr, insertTableRow as Ps, clearSlideBackground as Pt, setShapeReflection as Pu, getPresentationNotesText as Q, clearSlideAnimations as Qa, ChartTrendline as Qc, getShapesBounds as Qd, createPresentation as Qf, getUnusedSlideLayouts as Qi, getShapeTextAutoFit as Ql, resetSlidePlaceholderGeometry as Qn, setMediaPartBytes as Qo, Text3D as Qp, replaceTokensInPresentation as Qr, setTableStyleFlags as Qs, getSlideMasterBackgroundGradientFill as Qt, setShapePreset as Qu, findSlidesWithChartKind as R, setParagraphBullet as Ra, ChartDataLabels as Rc, findShapeByText as Rd, isShapePlaceholder as Rf, setSlideMasterName as Ri, getShapeHyperlink as Rl, cm as Rm, setShapeStyle as Rn, setShapeImageOpacity as Ro, GeomCommand as Rp, getSlideOutline as Rr, removeTableColumn as Rs, getSlideBackgroundImageBytes as Rt, clearShapeStroke as Ru, SlideSize as S, getParagraphBulletStyle as Sa, ComboChartSeries as Sc, getShapeFillColorResolved as Sd, getShapeId as Sf, setHandoutMasterPlaceholderIncluded as Si, getCommentSlide as Sl, ReadGradientStop as Sm, MediaFade as Sn, getShapeImageBytes as So, AnimationSpinDirection as Sp, findSlideByText as Sr, getTableCellSpan as Ss, getSlideTransition as St, Shape3D as Su, setSlideSize as T, getParagraphLevel as Ta, LineArea3DChartSpec as Tc, getShapeStroke as Td, getShapePlaceholderIdx as Tf, setNotesPageOrientation as Ti, getCommentsSortedByDate as Tl, ReflectionOptions as Tm, ShapeMedia as Tn, getShapeImageDuotone as To, TransitionEffect as Tp, findSlidesByText as Tr, getTableCellTextDirection as Ts, setSlideTransitionSound as Tt, clearShapeEffects as Tu, getAllImages as U, setParagraphTabs as Ua, ChartGrouping as Uc, findShapesByPreset as Ud, setShapeBounds as Uf, findSlideLayoutByPartName as Ui, resolveDrawingColor as Ul, CustomGeometryPathInput as Un, _internalPackageOf as Uo, BevelPreset as Up, getSlides as Ur, setTableCellClickAction as Us, getSlideBackgroundPatternFill as Ut, setShapeAdjustValues as Uu, getAllComments as V, setParagraphLineSpacing as Va, ChartErrorBarAmount as Vc, findShapesByKind as Vd, renameShape as Vf, findLayoutsWithPlaceholderType as Vi, setShapeRunFormat as Vl, mm as Vm, setShapeTextLanguage as Vn, PackagePartInfo as Vo, PathFillMode as Vp, getSlideTextLength as Vr, setTableCellAnchor as Vs, getSlideBackgroundImageIntrinsicSize as Vt, getShapeStrokeDash as Vu, getAllNotes as W, setParagraphTypography as Wa, ChartKind as Wc, findShapesByText as Wd, setShapeDescription as Wf, findSlideLayoutByType as Wi, resolveDrawingColorOpacity as Wl, setShapeCustomGeometry as Wn, compactPackage as Wo, CameraPreset as Wp, getSlidesByLayout as Wr, setTableCellFill as Ws, getSlideColorMapOverride as Wt, setShapeFill as Wu, getPresentationNotesLength as X, setSlidePlaceholders as Xa, ChartTextStyle as Xc, getShapeSlide as Xd, PresentationInput as Xf, getSlideLayoutUsageCountsByType as Xi, getShapeText3D as Xl, addSlidePlaceholder as Xn, listPackageParts as Xo, Rotation3D as Xp, replaceTextInSlide as Xr, setTableColumnWidth as Xs, getSlideLayoutShapes as Xt, setShapePatternFill as Xu, getPresentationHyperlinkCountsBySlide as Y, setSlideBody as Ya, ChartSeriesAxis as Yc, getShapeIndex as Yd, setSlideSections as Yf, getSlideLayoutUsageCounts as Yi, getShapeBodyPrEffective as Yl, addMissingSlidePlaceholders as Yn, getSlideMediaPartNames as Yo, ReadText3D as Yp, replaceTextInPresentation as Yr, setTableCellTextFormat as Ys, getSlideLayoutBackgroundShapes as Yt, setShapeNoStroke as Yu, getPresentationNotesLengthsBySlide as Z, setSlideTitle as Za, ChartTimeUnit as Zc, getShapeXmlString as Zd, PresentationSize as Zf, getSlideLayouts as Zi, getShapeTextAnchor as Zl, resetSlideLayout as Zn, readPackagePart as Zo, Scene3D as Zp, replaceTextInSlideNotes as Zr, setTableRowHeight as Zs, getSlideMasterBackground as Zt, setShapePosition as Zu, setSlideOutlineCollapsed as _, setShapeZIndex as _a, BubbleChartSeries as _c, setPresentationTheme as _d, getShapeBoundsResolved as _f, HandoutSlidesPerPage as _i, getCommentAuthors as _l, GradientFillOptions as _m, getShapeImageCompressionState as _n, findSlidesByLayoutType as _o, AnimationShape as _p, setCustomShows as _r, getTableCellFill as _s, setSlideNotesFormat as _t, setShapeTextFlat as _u, getShapeImageFillLayout as a, addSlideLine as aa, findChartsWithDataLabels as ac, setShapeStrokeCompound as ad, ShapeBounds as af, removeThumbnail as ai, ImageFit as al, ReadTextFormat as am, setSlideBackgroundImageFillLayout as an, AnimationTarget as ao, SlideShapeData as ap, addTitleSlide as ar, TableCellBorder as as, getSlidesWithComments as at, getShapeTextWrap as au, SLIDE_SIZE_16_9 as b, getParagraphBullet as ba, ChartSpec as bc, getShapeFill as bd, getShapeDescription as bf, getNotesMasterPlaceholders as bi, getCommentParent as bl, PatternPreset as bm, setShapePictureStyle as bn, getShapeImageBiLevelThreshold as bo, AnimationColor as bp, SlideOutlineEntry as br, getTableCellPosition as bs, TransitionSoundInput as bt, setShapeTextWrap as bu, DrawingGuide as c, addSlideTextBox as ca, getPresentationChartKindCounts as cc, setShapeStrokeSketch as cd, findShapesInRect as cf, ExtendedProperties as ci, clearAllSlideComments as cl, TextOutline as cm, ImageFillLayout as cn, ImageArtisticEffect as co, CommentStatus as cp, importSlide as cr, clearTableCellFill as cs, getSlidesWithNotes as ct, setShapeParagraphs as cu, getDrawingGuidesVisible as d, clearSlideShapes as da, getShapeChartSeriesNames as dc, PresentationFonts as dd, getMaxShapeId as df, getPresentationCreated as di, findCommentsAfter as dl, LineDash as dm, AnimationPatch as dn, ImageRecolorColor as do, SlideComment as dp, removeSlide as dr, getTableBackgroundEffective as ds, getVisibleSlides as dt, setShapeTextAnchor as du, clearSlideLayoutBackground as ea, splitTableCell as ec, setShapeSize as ed, getSlideShapes as ef, setSlideBackgroundGraphicsHidden as ei, ChartView3D as el, ParagraphAlignment as em, getSlideMasterBackgroundPatternFill as en, getShapeAnimation as eo, savePresentation as ep, addBlankSlide as er, validatePresentation as es, getSlideNotesLength as et, getShapeTextBodyRotationDeg as eu, getGridSpacing as f, copyShape as fa, getShapeChartSeriesValues as fc, PresentationFontsInput as fd, getMaxShapeIdInPresentation as ff, getPresentationModified as fi, findCommentsBefore as fl, LineEndSize as fm, moveSlideAnimation as fn, findFlippedShapes as fo, AnimationDirection as fp, reverseSlides as fr, getTableCell as fs, hasSlideNotes as ft, setShapeTextAutoFit as fu, setGridSpacing as g, sendShapeToBack as ga, Bar3DChartSpec as gc, setPresentationFonts as gd, getShapeBounds as gf, touchModified as gi, getCommentAuthor as gl, ReadLineFill as gm, ImageCompressionState as gn, findSlidesByLayoutPartName as go, AnimationOrientation as gp, getCustomShows as gr, getTableCellBorders as gs, setSlideNotes as gt, setShapeTextField as gu, setDrawingGuidesVisible as h, sendShapeBackward as ha, setChartSpec as hc, getPresentationTheme as hd, getShapeAltTitle as hf, setExtendedProperties as hi, findSlidesWithCommentsByAuthor as hl, LineSketch as hm, resetShapeVideoFormatting as hn, findSlidesByLayoutName as ho, AnimationOptions as hp, CustomShow as hr, getTableCellAppearanceEffective as hs, resolveSlideNotesTextColor as ht, setShapeTextDirection as hu, setSlideShowProperties as i, addSlideImage as ia, findChartsBySeriesName as ic, setShapeStrokeCap as id, translateShapes as if, getThumbnail as ii, setShapeClickAction as il, ReadTextFill as im, setSlideBackgroundImage as in, AnimationStart as io, SlideLayoutData as ip, addSlideAt as ir, TableCellAppearanceEffective as is, getSlidesWithCharts as it, getShapeTextMargins as iu, PresentationTableEntry as j, getShapeRunClickAction as ja, ScatterChartSeries as jc, getShapeStrokeGradient as jd, getShapeText as jf, getSlideMasterPlaceholders as ji, removeSlideComment as jl, ColorTransform as jm, getShapeMediaPlayback as jn, hasShapeImage as jo, NotesMasterPlaceholderType as jp, getPresentationTextLengthsBySlide as jr, getTableStyleFlags as js, SlideBackground as jt, setShape3D as ju, PresentationImageEntry as k, getShapeParagraphCount as ka, PlainChartSeries as kc, getShapeStrokeCompound as kd, getShapeRotation as kf, getSlideMasterLayouts as ki, getSlideCommentAuthors as kl, SchemeColorToken as km, findShapesWithMedia as kn, getShapeImageOpacity as ko, LayoutPlaceholderKind as kp, getPresentationText as kr, getTableRowHeights as ks, getSlideMasterBackgroundStyles as kt, getShapeEffects as ku, getCollapsedOutlineSlides as l, bringShapeForward as la, getShapeChartCategories as lc, getShapeGradientFill as ld, getGroupChildren as lf, getCoreProperties as li, clearSlideComments as ll, toWritableTextFormat as lm, ImageTileAlignment as ln, ImageCrop as lo, CommentAuthor as lp, mergePresentations as lr, getPresentationTableCountsBySlide as ls, getSlidesWithOverlap as lt, setShapeText as lu, setDrawingGuides as m, removeShape as ma, getSlideCharts as mc, getPresentationFonts as md, getShapeAdjustValues as mf, setCoreProperties as mi, findCommentsByText as ml, LineFill as mm, updateSlideAnimation as mn, findShapesOutsideCanvas as mo, AnimationInOut as mp, swapSlides as mr, getTableCellAnchor as ms, replaceHyperlink as mt, setShapeTextColumns as mu, SlideShowProperties as n, setSlideLayoutName as na, addSlideChart as nc, setShapeStroke as nd, replaceTokensInSlide as nf, slideHasAnimations as ni, ShapeClickAction as nl, ParagraphSpec as nm, setSlideBackground as nn, setShapeAnimation as no, SlideCommentData as np, addSectionHeaderSlide as nr, ValidationIssue as ns, getSlideNotesParagraphEndFormat as nt, getShapeTextDirection as nu, getShapeImageIntrinsicSize as o, addSlideShape as oa, findChartsWithTrendlines as oc, setShapeStrokeDash as od, centerShapeOnSlide as of, setThumbnail as oi, setShapeImage as ol, TextFill as om, setSlideBackgroundImageOpacity as on, AnimationValueAfterEnd as oo, TableCellData as op, duplicateSlide as or, TableCellBorders as os, getSlidesWithHyperlinks as ot, setShapeAlignment as ou, getSnapToGrid as p, getShapeZIndex as pa, getShapeChartSpec as pc, PresentationTheme as pd, getPresetGeometry as pf, incrementRevision as pi, findCommentsByAuthor as pl, LineEndType as pm, removeSlideAnimation as pn, findOverlappingShapePairs as po, AnimationEffect as pp, sortSlides as pr, getTableCellAlignment as ps, removeSlideNotes as pt, setShapeTextBodyRotationDeg as pu, getEmptySlides as q, getSlideBody as qa, ChartSecondaryValueAxis as qc, findSlidePlaceholderByIdx as qd, SlideSection as qf, getSlideLayoutPlaceholders as qi, TextWrap as ql, setShapeAspectRatioLocked as qn, getPackageSize as qo, LightRigType as qp, isSlideLayoutBackgroundGraphicsHidden as qr, setTableCellText as qs, getSlideLayoutBackgroundImageBytes as qt, setShapeImageFill as qu, getSlideShowProperties as r, setSlideLayoutPlaceholderBounds as ra, findChartByKind as rc, setShapeStrokeArrow as rd, setSlideLayout as rf, PresentationThumbnail as ri, getShapeClickAction as rl, RunSpec as rm, setSlideBackgroundGradientFill as rn, AnimationSequenceKind as ro, SlideData as rp, addSlide as rr, TableBackgroundEffective as rs, getSlideNotesTextFormats as rt, getShapeTextFlat as ru, setShapeImageFillLayout as s, addSlideTable as sa, getPresentationChartCountsBySlide as sc, setShapeStrokeJoin as sd, findShapesAtPoint as sf, CoreProperties as si, addSlideComment as sl, TextFormat as sm, setSlideBackgroundPatternFill as sn, SlideAnimationStep as so, COMMENT_STATUSES as sp, duplicateSlideAt as sr, TableCellParagraph as ss, getSlidesWithImages as st, setShapeBulletStyle as su, VERSION as t, setSlideLayoutBackground as ta, SlideChartData as tc, setShapeSlideBackgroundFill as td, getSlidesWithEmptyPlaceholders as tf, setSlideHidden as ti, ReadChartSpec as tl, ParagraphAlignmentToken as tm, getSlideMasterShapes as tn, getSlideAnimations as to, PresentationData as tp, addContentSlide as tr, IssueSeverity as ts, getSlideNotesLineBreaks as tt, getShapeTextColumns as tu, getDrawingGuides as u, bringShapeToFront as ua, getShapeChartKind as uc, getShapeGradientFillEffective as ud, getGroupTransform as uf, getExtendedProperties as ui, findCommentAuthorByName as ul, ArrowOptions as um, ImageTileFlip as un, ImageRecolor as uo, CommentPosition as up, moveSlide as ur, getSlideTables as us, getSlidesWithTables as ut, setShapeText3D as uu, setSnapToGrid as v, ShapeParagraphElement as va, BubbleChartSpec as vc, ShapeFill as vd, getShapeCenter as vf, getHandoutMasterPlaceholders as vi, getCommentDate as vl, GradientStop as vm, getShapePictureStyle as vn, getPresentationImageCountsBySlide as vo, AnimationStartCondition as vp, AllShapesEntry as vr, getTableCellMargins as vs, transformSlideNotesCase as vt, setShapeTextFormat as vu, getSlideSize as w, getParagraphIndent as wa, DoughnutChartSpec as wc, getShapeFillOpacity as wd, getShapeName as wf, setNotesMasterPlaceholderIncluded as wi, getCommentText as wl, InnerShadowOptions as wm, MediaTrim as wn, getShapeImageCrop as wo, SlideTransition as wp, findSlidesByNotes as wr, getTableCellText3D as ws, setSlideTransition as wt, ShapeEffectAny as wu, SLIDE_SIZE_4_3 as x, getParagraphBulletImageBytes as xa, ChartVariant as xc, getShapeFillColor as xd, getShapeFlip as xf, getNotesPageSize as xi, getCommentPosition as xl, ReadGradientFill as xm, MediaBookmark as xn, getShapeImageBrightness as xo, AnimationScaleDirection as xp, findSlideByPartName as xr, getTableCellRunFormatEffective as xs, clearSlideTransition as xt, TextCase as xu, SLIDE_SIZE_16_10 as y, getParagraphAlignment as ya, CandlestickChartSpec as yc, ShapeStroke as yd, getShapeCustomGeometry as yf, getHandoutSlidesPerPage as yi, getCommentFormat as yl, PatternFillOptions as ym, setShapeImageCompressionState as yn, getShapeImageArtisticEffect as yo, AnimationTextBuild as yp, SlideInfo as yr, getTableCellParagraphs as ys, TransitionSound as yt, setShapeTextMargins as yu, findSlidesWithChartTrendlines as z, setParagraphIndent as za, ChartDataTable as zc, findShapeInPresentation as zd, isShapeTextBox as zf, setSlideMasterPlaceholderIncluded as zi, getShapeRunFormatEffective as zl, emu as zm, getParagraphElementLanguages as zn, setShapeImageRecolor as zo, GeomPath as zp, getSlidePartName as zr, removeTableRow as zs, getSlideBackgroundImageCrop as zt, getShapePatternFill as zu };
//# sourceMappingURL=index-WCK8v0WT.d.ts.map