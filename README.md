# TesoroRGB

TesoroRGB is a Library for modifying the lighting profiles of Tesoro RGB Keyboards.

## Supported Devices
The following devices have been tested and confirmed to be compatible. If you own a device that isn't listed, please try it out and let me know the results.

- Tesoro Gram Spectrum

## Example
```
// Create Keyboard object
Keyboard kb = new Keyboard();

// Find the Keyboard and establish connection
kb.Initialize();

// Set the keyboard to display PC mode
kb.SetProfile(TesoroProfile.PC);
Thread.Sleep(100);

// Set the lighting mode to custom colours with a breathing effect
kb.SetLightingMode(LightingMode.SpectrumColors, SpectrumMode.SpectrumBreathing, TesoroProfile.PC);
Thread.Sleep(100);

// Set the keys to display the contents of an image file
kb.SetKeysColor("test.png", TesoroProfile.PC, false);

// Set the keys to display the contents of a Bitmap object
kb.SetKeysColor(bitmap, TesoroProfile.PC, false);

// Set individual key color
kb.SetKeyColor(TesoroLedID.Escape, 255, 0, 0, TesoroProfile.PC);

// Save Layout
kb.SaveSpectrumColors(TesoroProfile.PC);

// Release handle, close filestreams
kb.Uninitialize();
```
