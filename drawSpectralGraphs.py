from matplotlib import pyplot as plt
import numpy as np

def drawHorGraphs(sd1, sds):
    seconds = np.arange(0,10,0.01)
    T_B = sd1 / sds
    T_A = 0.2 * T_B
    T_L = 6.0
    valuesSae = []
    for T in seconds:
        if 0 <= T <= T_A:
            Sae = (0.4 + 0.6 * (T / T_A)) * sds
            valuesSae.append(Sae)
        elif T_A < T <= T_B:
            Sae = sds
            valuesSae.append(Sae)
        elif T_B < T <= T_L:
            Sae = sd1 / T
            valuesSae.append(Sae)
        else:
            Sae = sd1 * (T_L / pow(T, 2))
            valuesSae.append(Sae)
    plt.figure(figsize=(3, 2.5))
    plt.plot(seconds, valuesSae, "blue")
    plt.xlabel("periyot")
    plt.ylabel("yatay elastik tasarım spektral ivmeleri")
    plt.savefig("horiSpect.png", dpi=100)
    plt.close()
def drawHVerGraphs(sd1, sds):
    seconds = np.arange(0, 3, 0.01)
    T_BD = (sd1 / sds) / 3
    T_AD = (0.2 * T_BD) / 3
    T_LD = 3.0
    valuesSaeD = []
    for T in seconds:
        if 0 <= T <= T_AD:
            SaeD = (0.32 + 0.48 * (T / T_AD)) * sds
            valuesSaeD.append(SaeD)
        elif T_AD < T <= T_BD:
            SaeD = 0.8 * sds
            valuesSaeD.append(SaeD)
        elif T_BD < T <= T_LD:
            SaeD = (0.8 * sds) * (T_BD / T)
            valuesSaeD.append(SaeD)
    plt.figure(figsize=(3, 2.5))
    plt.plot(seconds, valuesSaeD, "red")
    plt.xlabel("periyot")
    plt.ylabel("düşey elastik tasarım spektral ivmeleri")
    plt.savefig("vertSpect.png", dpi=100)
    plt.close()
