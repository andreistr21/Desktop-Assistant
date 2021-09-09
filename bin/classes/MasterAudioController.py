from ctypes import cast, POINTER

from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class MasterAudioController(object):
    def __init__(self):
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        # noinspection PyTypeChecker
        self.pc_volume = cast(interface, POINTER(IAudioEndpointVolume))

    def GetVolumeRange(self):
        return self.pc_volume.GetVolumeRange()

    def SetVolume(self, decibels):
        self.pc_volume.SetMasterVolumeLevel(decibels, None)

    def DecreaseVolume(self):
        pass