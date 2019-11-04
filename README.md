# Driver-MoxaIOLogik_E1200Series

PACScript Driver to communicate with MOXA ioLogik E1200 series devices. 

- [Release Page](https://github.com/DENSO-2DLab/RC8_Driver-MoxaIOLogik_E1200Series/releases)

### Contents

It is being more common to have a need for the DENSO robot controller to act as a master device and be at the center of the control sequence. This requires having capability to expand IO and if possible eliminate the need for a PLC. It is possible to make use of the Modbus provider feature of the Denso controller to expand the IO capabilities by using a remote IO block like the Moxa ioLogik E1200 series. 

Moxa ioLogik is compact and can provide additional Digital Inputs and Output as well as Analog Inputs. Since it uses Modbus TCP it is possible to connect to a multiple of devices.

#### DENSO System Requirements

- **Robot Controller Type**: COBOTTA, RC8, RC8A
- **Minimum Firmware Version**: 1.12.*
- **Extensions Necessary**: ModBus Provider Extension

### Contributing 

Feel free to contribute by fixing any bugs and adding new features to this repository. 
Once you are ready to share open a new pull request or generate an Issue. 
- [Submit an Issue](https://github.com/DENSO-2DLab/RC8_Driver-MoxaIOLogik_E1200Series/pulls)

### License 

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

### Release Notes 

**v1.0.0** 
- Initial Release
