import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np

# 1. Complete AMBER dictionary with all keys required by the script
AMBER = {
    "window_bg": "#100b05",
    "amber_primary": "#ffb000",
    "amber_muted": "#805b18",
    "text_primary": "#ffc24b",
    "text_secondary": "#c99432",
}

# 2. Define the Custom Colormap
colors = [AMBER["window_bg"], AMBER["amber_muted"], AMBER["amber_primary"]]
gui_amber_cmap = mcolors.LinearSegmentedColormap.from_list("gui_amber_theme", colors)

# 3. Generate Synthetic Image Data
synthetic_img = np.zeros((500, 500), dtype=float)

# Left third: Linear gradient
synthetic_img[:, :166] = np.linspace(0, 1, 500)[:, np.newaxis]

# Middle third: Stripes
stripe_h = 125
synthetic_img[0:stripe_h, 166:332] = 0.25
synthetic_img[stripe_h:2*stripe_h, 166:332] = 0.50
synthetic_img[2*stripe_h:3*stripe_h, 166:332] = 0.75
synthetic_img[3*stripe_h:, 166:332] = 1.00

# Right third: Radial gradient (Correctly sliced)
Y, X = np.ogrid[:500, :500]
dist = np.sqrt((X - 415)**2 + (Y - 250)**2)
radial_pattern = np.clip(1.0 - (dist / 150), 0, 1)
synthetic_img[:, 332:] = radial_pattern[:, 332:]

# 4. Display
fig = plt.figure(figsize=(10, 7))
fig.patch.set_facecolor(AMBER["window_bg"]) # Sets background outside the plot
ax = plt.gca()
ax.set_facecolor(AMBER["window_bg"])

img_plot = ax.imshow(synthetic_img, cmap=gui_amber_cmap, vmin=0, vmax=1)

# Colorbar
cbar = plt.colorbar(img_plot, ticks=[0, 0.25, 0.5, 0.75, 1.0])
cbar.ax.set_yticklabels(['BG', '25%', '50%', '75%', 'Primary'], color=AMBER["text_secondary"])
cbar.set_label('Intensity', color=AMBER["text_primary"])

plt.title("Synthetic Amber Test", color=AMBER["text_primary"])
plt.axis('off')
plt.show()