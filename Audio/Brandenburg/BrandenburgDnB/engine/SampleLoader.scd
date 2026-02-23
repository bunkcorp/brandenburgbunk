/*
  Sample Loader Module
  Loads WAV/AIFF, optional slice points (1-bar, 1/4 note, or manual).
  Target BPM used for time-stretch or slice grid (default 172).
*/
(
~loadBrandenburg = { |path, bpm = 172, numSlices = 16|
  Buffer.read(s, path, 0, -1, { |buf|
    var len = buf.numFrames;
    var sliceFrames = (len / numSlices).floor.asInteger;
    var points = (0..numSlices).collect(_ * sliceFrames).min(len);
    "Brandenburg loaded: % frames, % slices at ~% frames each, BPM %".format(len, numSlices, sliceFrames, bpm).postln;
    ~brandenburgBuf = buf;
    ~brandenburgSlices = points;
    ~brandenburgBpm = bpm;
    ~brandenburgNumSlices = numSlices;
  });
};

~getSliceFrames = { |sliceIndex|
  var slices = ~brandenburgSlices;
  if (slices.isNil) { [0, 44100] } {
    var i = sliceIndex.wrap(0, slices.size - 2);
    [slices[i], slices[i + 1] - slices[i]]
  }
};
);
