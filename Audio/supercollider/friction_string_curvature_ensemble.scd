// Friction string "Curvature-style" ensemble:
// one wheel concept, many strings with different resonances.
(
~configureFrictionServer = { |outDevice|
    s.options.numInputBusChannels = 0;
    s.options.outDevice = outDevice;
    s.options.blockSize = 512;
    ("friction_string_curvature_ensemble: trying output -> " ++ outDevice).postln;
};

~startFrictionEnsemble = {
    SynthDef(\frictionString, {
        |out = 0, freq = 160, pressure = 0.75, bowRate = 0.12, amp = 0.16,
         bodyTilt = 1.0, pan = 0.0, wheelPhase = 0.0|

        var bowNoise;
        var wheel;
        var excite;
        var string;
        var body;
        var drift;
        var drift2;
        var feedback;
        var reverb;
        var freqSmooth;
        var scrape;
        var dryStereo;

        drift = LFNoise2.kr(0.13).range(0.97, 1.03);
        drift2 = SinOsc.kr(0.037).range(0.985, 1.015);
        freqSmooth = Lag.kr(freq, 1.8);

        // "One wheel, many strings" illusion: shared rate region, offset phase.
        wheel = LFSaw.kr(bowRate + (LFNoise1.kr(0.25) * 0.03), wheelPhase).range(0, 1);

        bowNoise = PinkNoise.ar * pressure;
        bowNoise = bowNoise * (wheel + 0.08);
        bowNoise = LeakDC.ar(bowNoise);

        excite = tanh(bowNoise * 5.5);
        excite = excite + (SinOsc.ar(freqSmooth * 0.5) * 0.015);

        string = CombL.ar(
            excite,
            0.25,
            (1 / (freqSmooth * drift * drift2)).clip(0.001, 0.22),
            11
        );

        body =
            BPF.ar(string, freqSmooth * (1.00 * bodyTilt), 0.22) +
            BPF.ar(string, freqSmooth * (1.61 * bodyTilt), 0.21) +
            BPF.ar(string, freqSmooth * (2.32 * bodyTilt), 0.25) +
            BPF.ar(string, freqSmooth * (3.73 * bodyTilt), 0.30);
        body = body + (LPF.ar(string, 3200) * 0.25);

        scrape = HPF.ar(PinkNoise.ar * pressure, 3800) * 0.03;
        body = body + scrape;

        feedback = LocalIn.ar(1);
        body = body + (feedback * 0.18);
        LocalOut.ar(Limiter.ar(LeakDC.ar(Mix(body)), 0.9));

        reverb = FreeVerb2.ar(body, body, 0.58, 0.9, 0.35);
        dryStereo = Splay.ar([body, DelayC.ar(body, 0.03, 0.009)], spread: 0.2);
        dryStereo = Balance2.ar(dryStereo[0], dryStereo[1], pan);

        Out.ar(out, Limiter.ar((dryStereo + reverb) * amp, 0.95));
    }).add;

    s.sync;
    "friction_string_curvature_ensemble: server booted".postln;

    // Ensemble settings.
    ~ensembleSize = 12;
    ~baseScale = [48, 50, 52, 55, 57, 59, 62, 64]; // airy modal center
    ~voices = Array.fill(~ensembleSize, { |i|
        var n = ~baseScale.wrapAt(i) + [0, 12].wchoose([0.7, 0.3]);
        Synth(\frictionString, [
            \freq, (n + rrand(-0.1, 0.1)).midicps,
            \pressure, rrand(0.50, 0.78),
            \bowRate, rrand(0.05, 0.15),
            \amp, 0.12 + (0.06.rand),
            \bodyTilt, exprand(0.92, 1.08),
            \pan, ((i / (~ensembleSize - 1)) * 2 - 1) * 0.8,
            \wheelPhase, i / ~ensembleSize
        ]);
    });

    // Slow flocking motion across all voices.
    if(~ensembleRoutine.notNil) { ~ensembleRoutine.stop; };
    ~ensembleRoutine = Routine({
        var t = 0;
        loop {
            ~voices.do { |v, i|
                var n = ~baseScale.wrapAt(i + t) + [-12, 0, 7, 12].wchoose([0.10, 0.65, 0.18, 0.07]);
                n = n.clip(45, 81);
                v.set(
                    \freq, (n + rrand(-0.12, 0.12)).midicps,
                    \pressure, rrand(0.46, 0.80),
                    \bowRate, rrand(0.04, 0.16),
                    \bodyTilt, exprand(0.90, 1.10)
                );
            };
            t = t + 1;
            rrand(2.8, 5.8).wait;
        }
    }).play(AppClock);

    "friction_string_curvature_ensemble: voices running".postln;

    // Stop helpers:
    // ~ensembleRoutine.stop;
    // ~voices.do(_.free);
};

~bootWithFallback = { |primaryOut, fallbackOut|
    ~configureFrictionServer.(primaryOut);
    s.waitForBoot(
        { ~startFrictionEnsemble.(); },
        100,
        {
            ("friction_string_curvature_ensemble: failed on -> " ++ primaryOut).warn;
            if(fallbackOut.notNil) {
                AppClock.sched(0.5, {
                    ~bootWithFallback.(fallbackOut, nil);
                    nil
                });
            };
        }
    );
};

~bootWithFallback.("Kevin’s AirPods P", "MacBook Pro Speakers");
)
