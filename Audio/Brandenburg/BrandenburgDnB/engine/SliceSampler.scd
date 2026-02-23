/*
  Slice Sampler – plays back slice from buffer with speed/pitch, optional reverse.
  Out to a bus for feeding into SP1200/S760 and FX.
*/
(
SynthDef(\brandenburg_slice, { |bufnum, startFrame = 0, numFrames = 4410, speed = 1, reverse = 0, outBus = 0, amp = 0.8, gate = 1|
  var phase = Phasor.ar(0, BufRateScale.kr(bufnum) * speed * (1 - (reverse * 2)), startFrame, startFrame + numFrames);
  var snd = BufRd.ar(2, bufnum, phase, 1, 4);
  snd = snd * EnvGen.kr(Env.asr(0.005, 1, 0.1), gate, doneAction: 2) * amp;
  Out.ar(outBus, snd);
}).add;
);
