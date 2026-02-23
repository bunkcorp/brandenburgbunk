-- Brandenburg DnB – TidalCycles pattern sketch
-- Send OSC to SuperCollider Brandenburg engine (run main.scd / run_brandenburg.scd first).
-- Uses OSC to /brandenburg/play_slice and /brandenburg/dsp_mode etc.
-- For SuperDirt + samples, use standard d1/d2; for Brandenburg engine use the OSC pattern below.

import Sound.Tidal.OSC

-- Example: if you have a separate OSC target for Brandenburg engine at 127.0.0.1:57120
-- tidalSend (sock) $ OSC "/brandenburg/play_slice" [OSC_I 0, OSC_F 1, OSC_I 0, OSC_F 0.5]

-- SuperDirt-based pattern (works with your existing SuperDirt on 57120):
-- Brandenburg-style melody + DnB break (run in ghci with BootTidal or :script this file)

-- d1 $ stack [
--   sound "bd sn bd sn",
--   sound "hh*8",
--   sound "sn(3,8)"
-- ] # speed 1 # n (run 8)

-- d2 $ slow 2 $ sound "superpiano" # n "b4 fs4 b4 d5" # sustain 0.3

-- To drive the Brandenburg engine from Tidal you’d add a custom OSC stream in BootTidal
-- that sends /brandenburg/play_slice and /brandenburg/dsp_mode (0 or 1) from pattern values.
