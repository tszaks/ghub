import Vision
import AppKit

guard CommandLine.arguments.count >= 2 else {
    FileHandle.standardError.write(Data("usage: ocr <image_path>\n".utf8))
    exit(2)
}

let path = CommandLine.arguments[1]

guard let image = NSImage(contentsOfFile: path),
      let cgImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    FileHandle.standardError.write(Data("failed to load image at \(path)\n".utf8))
    exit(1)
}

let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.usesLanguageCorrection = true

let handler = VNImageRequestHandler(cgImage: cgImage, options: [:])

do {
    try handler.perform([request])
} catch {
    FileHandle.standardError.write(Data("vision request failed: \(error)\n".utf8))
    exit(1)
}

let strings = (request.results ?? []).compactMap { observation in
    observation.topCandidates(1).first?.string
}

print(strings.joined(separator: "\n"))
