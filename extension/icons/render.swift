import AppKit
let svgPath = CommandLine.arguments[1], outDir = CommandLine.arguments[2]
guard let img = NSImage(contentsOfFile: svgPath) else { print("cannot load svg"); exit(1) }
for s in [16, 32, 48, 128] {
  let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: s, pixelsHigh: s, bitsPerSample: 8, samplesPerPixel: 4,
    hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
  rep.size = NSSize(width: s, height: s)
  NSGraphicsContext.saveGraphicsState()
  NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)
  NSGraphicsContext.current?.imageInterpolation = .high
  // The Web Store shows the 128 icon and asks for 96×96 artwork inside 16px of
  // transparent padding; the toolbar sizes fill their square.
  let inset = s == 128 ? 16 : 0
  img.draw(in: NSRect(x: inset, y: inset, width: s - 2 * inset, height: s - 2 * inset), from: .zero, operation: .sourceOver, fraction: 1)
  NSGraphicsContext.restoreGraphicsState()
  try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: "\(outDir)/icon\(s).png"))
}
print("ok")
