// Amen loop — buffer playback with wobble filter and rhythm-sliced position
// Set ~amenPath to your Amen WAV (e.g. ~/Downloads/amen.wav). Requires server booted.

var amenPath = "/Users/kevinwoods/Downloads/amen-break-frenzy-jazzy_174bpm_G#_major.wav";

s.waitForBoot {

	~amenPath = amenPath;
	~amen = Buffer.read(s, ~amenPath);

	s.sync;

	SynthDef(\loop, { |out = 0, bufnum = 0, gate = 1, pos = 0, speed = 1, freq = 0, endfreq = 0.001, sustain = 1, wobble = 3, boost = 1|
		var player, env, amp;
		freq = XLine.ar(freq, endfreq, sustain / 4);
		freq = freq.cpsmidi + (LFNoise2.ar(3).range(-1, 1) * (1/12));
		freq = freq.midicps;
		env = Env.adsr(0.0001, 0.01, 1, sustain / 8, 1);
		amp = EnvGen.kr(env, gate, doneAction: 2);
		player = PlayBuf.ar(2, bufnum, BufRateScale.kr(bufnum) * speed, Impulse.ar(freq), BufFrames.kr(bufnum) * pos + Rand(0, 20), 1, 2) * boost;
		player = RLPF.ar(player, SinOsc.ar(wobble / sustain).range(20000, 80), XLine.ar(0.2, 0.9, sustain)) * boost;
		Out.ar(out, player * amp);
	}).add;

	SynthDef(\compressor, {
		var in, compressed;
		in = In.ar(0, 2);
		compressed = Compander.ar(in, in, 0.1, 1, 1/8, 0.002, 0.01);
		ReplaceOut.ar(0, compressed * 4);
	}).add;

	s.sync;

	TempoClock.default.tempo = 16 / ~amen.duration;

	// Rhythm: if .convertRhythm is undefined (custom/old), use one of the Pseq durs below instead
	~rhythm = Pseq([8, [2, [1, [1, 1, [1, 1!3], [1, 1!4]]], [1, 1!4]]].convertRhythm, 1);
	// ~rhythm = Pseq([8, [[2, [2, 1, [1, 1!12]]], [1, [1, 1, [1, 1!3], [1, 1!4]]], [1, 1!4]]].convertRhythm, 1);
	// ~rhythm = Pseq([8, 1!16].convertRhythm, 1);
	// Fallback if no convertRhythm: ~rhythm = Pseq(0.25 ! 16, 1);  // 16 16ths

	~amenPlayer = Pbind(
		\instrument, \loop,
		\bufnum, ~amen.bufnum,
		\out, 0,
		\freq, Pstep([43.midicps, 0, 0, 0], 2, inf),
		\endfreq, Pstep([40.midicps / 2, 0, 0, 0], 2, inf),
		\wobble, Pstep([Prand([1/2, 2, 1, 1/3]), 0, 0, 0], 2, inf),
		\pos, Pstep(Array.series(8, 0, 1), 1, inf) / 8,
		\legato, 1,
		\sustain, Pkey(\dur),
		\dur, Pn(Plazy { ~rhythm })
	).play;

	~comp = Synth(\compressor, addAction: \addToTail);

	"Amen loop playing. Stop: ~amenPlayer.stop; ~comp.free; ~amen.free;".postln;
};

// Stop and free:
// ~amenPlayer.stop; ~comp.free; ~amen.free;
