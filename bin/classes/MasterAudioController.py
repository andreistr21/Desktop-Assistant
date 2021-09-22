from ctypes import cast, POINTER

from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class MasterAudioController(object):
    def __init__(self):
        devices = AudioUtilities.GetSpeakers()
        # noinspection PyProtectedMember
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        # noinspection PyTypeChecker
        self.pc_volume = cast(interface, POINTER(IAudioEndpointVolume))

    def GetVolumeRange(self):
        return self.pc_volume.GetVolumeRange()

    def SetVolumeScalar(self, percentages, is_decimal=False):
        if not is_decimal:
            percentages /= 100
        self.pc_volume.SetMasterVolumeLevelScalar(percentages, None)

    def ChangeVolume(self, decibels):
        self.pc_volume.SetMasterVolumeLevel(decibels, None)

    def GetMasterVolume(self):
        return self.pc_volume.GetMasterVolumeLevelScalar()

    def DecreaseVolume(self):
        pass
