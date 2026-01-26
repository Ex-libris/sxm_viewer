import numpy as np
from scipy import ndimage
from skimage import registration
from skimage.feature import match_template
import matplotlib.pyplot as plt
from pathlib import Path
from typing import List, Tuple, Optional
import cv2


class MicroscopyAligner:
    """Align and crop time-lapse microscopy images with drift correction."""
    
    def __init__(self, images: List[np.ndarray]):
        """
        Initialize with a list of images (as numpy arrays).
        
        Args:
            images: List of 2D or 3D numpy arrays (grayscale or RGB)
        """
        self.images = images
        self.n_images = len(images)
        self.shifts = None
        self.aligned_images = None
        self.cropped_images = None
        self.crop_bounds = None
        
    def compute_shifts(self, reference_idx: int = 0, upsample_factor: int = 10) -> np.ndarray:
        """
        Compute shift between each image and reference using phase correlation.
        
        Args:
            reference_idx: Index of reference image (default: first image)
            upsample_factor: Precision of subpixel registration
            
        Returns:
            Array of shifts (n_images, 2) with (dy, dx) for each image
        """
        reference = self._to_grayscale(self.images[reference_idx])
        shifts = np.zeros((self.n_images, 2))
        
        for i, img in enumerate(self.images):
            if i == reference_idx:
                continue
                
            target = self._to_grayscale(img)
            
            # Phase correlation for subpixel registration
            shift, error, diffphase = registration.phase_cross_correlation(
                reference, target, 
                upsample_factor=upsample_factor,
                normalization=None
            )
            shifts[i] = shift
            
        self.shifts = shifts
        return shifts
    
    def align_images(self, order: int = 3) -> List[np.ndarray]:
        """
        Align images using computed shifts.
        
        Args:
            order: Interpolation order (0=nearest, 1=bilinear, 3=cubic)
            
        Returns:
            List of aligned images
        """
        if self.shifts is None:
            self.compute_shifts()
        
        aligned = []
        for i, img in enumerate(self.images):
            # Apply negative shift to align
            shift = -self.shifts[i]
            
            if img.ndim == 2:
                # Grayscale
                aligned_img = ndimage.shift(img, shift, order=order, mode='constant', cval=0)
            else:
                # RGB - shift each channel
                aligned_img = np.zeros_like(img)
                for c in range(img.shape[2]):
                    aligned_img[:, :, c] = ndimage.shift(
                        img[:, :, c], shift, order=order, mode='constant', cval=0
                    )
            
            aligned.append(aligned_img)
        
        self.aligned_images = aligned
        return aligned
    
    def compute_crop_bounds(self, padding: int = 0) -> Tuple[int, int, int, int]:
        """
        Compute the maximum common crop area across all aligned images.
        
        Args:
            padding: Additional pixels to crop from edges (safety margin)
            
        Returns:
            Tuple of (top, bottom, left, right) crop bounds
        """
        if self.shifts is None:
            raise ValueError("Must compute shifts first")
        
        h, w = self.images[0].shape[:2]
        
        # Find cumulative shift bounds
        dy_shifts = self.shifts[:, 0]
        dx_shifts = self.shifts[:, 1]
        
        # Maximum positive shifts define how much to crop from top/left
        max_dy_pos = max(0, np.max(dy_shifts))
        max_dx_pos = max(0, np.max(dx_shifts))
        
        # Maximum negative shifts define how much to crop from bottom/right
        max_dy_neg = max(0, -np.min(dy_shifts))
        max_dx_neg = max(0, -np.min(dx_shifts))
        
        # Add padding and convert to integers
        top = int(np.ceil(max_dy_pos)) + padding
        bottom = h - int(np.ceil(max_dy_neg)) - padding
        left = int(np.ceil(max_dx_pos)) + padding
        right = w - int(np.ceil(max_dx_neg)) - padding
        
        # Ensure bounds are valid
        top = max(0, min(top, h))
        bottom = max(top + 1, min(bottom, h))
        left = max(0, min(left, w))
        right = max(left + 1, min(right, w))
        
        self.crop_bounds = (top, bottom, left, right)
        return self.crop_bounds
    
    def crop_to_common_area(self, images: Optional[List[np.ndarray]] = None) -> List[np.ndarray]:
        """
        Crop images to common area.
        
        Args:
            images: Images to crop (default: aligned images)
            
        Returns:
            List of cropped images
        """
        if images is None:
            if self.aligned_images is None:
                raise ValueError("Must align images first")
            images = self.aligned_images
        
        if self.crop_bounds is None:
            self.compute_crop_bounds()
        
        top, bottom, left, right = self.crop_bounds
        
        cropped = []
        for img in images:
            if img.ndim == 2:
                cropped.append(img[top:bottom, left:right])
            else:
                cropped.append(img[top:bottom, left:right, :])
        
        self.cropped_images = cropped
        return cropped
    
    def process_all(self, reference_idx: int = 0, padding: int = 5, 
                    order: int = 3) -> List[np.ndarray]:
        """
        Complete pipeline: align and crop all images.
        
        Args:
            reference_idx: Reference image index
            padding: Safety margin in pixels
            order: Interpolation order
            
        Returns:
            List of aligned and cropped images
        """
        print(f"Processing {self.n_images} images...")
        
        print("1. Computing shifts...")
        self.compute_shifts(reference_idx=reference_idx)
        print(f"   Max shift: {np.max(np.abs(self.shifts)):.2f} pixels")
        
        print("2. Aligning images...")
        self.align_images(order=order)
        
        print("3. Computing crop bounds...")
        self.compute_crop_bounds(padding=padding)
        top, bottom, left, right = self.crop_bounds
        print(f"   Crop region: [{top}:{bottom}, {left}:{right}]")
        print(f"   Output size: {bottom-top} x {right-left}")
        
        print("4. Cropping to common area...")
        self.crop_to_common_area()
        
        print("Done!")
        return self.cropped_images
    
    def visualize_results(self, indices: Optional[List[int]] = None, 
                         figsize: Tuple[int, int] = (15, 10)):
        """
        Visualize original, aligned, and cropped images.
        
        Args:
            indices: Image indices to show (default: first, middle, last)
            figsize: Figure size
        """
        if indices is None:
            indices = [0, self.n_images // 2, self.n_images - 1]
        
        n_show = len(indices)
        fig, axes = plt.subplots(3, n_show, figsize=figsize)
        
        if n_show == 1:
            axes = axes.reshape(-1, 1)
        
        for col, idx in enumerate(indices):
            # Original
            axes[0, col].imshow(self.images[idx], cmap='gray')
            axes[0, col].set_title(f'Original (t={idx})')
            axes[0, col].axis('off')
            
            # Aligned
            if self.aligned_images is not None:
                axes[1, col].imshow(self.aligned_images[idx], cmap='gray')
                axes[1, col].set_title(f'Aligned (shift={self.shifts[idx]})')
                axes[1, col].axis('off')
            
            # Cropped
            if self.cropped_images is not None:
                axes[2, col].imshow(self.cropped_images[idx], cmap='gray')
                axes[2, col].set_title(f'Cropped')
                axes[2, col].axis('off')
        
        plt.tight_layout()
        plt.show()
    
    @staticmethod
    def _to_grayscale(img: np.ndarray) -> np.ndarray:
        """Convert image to grayscale if needed."""
        if img.ndim == 3:
            # RGB to grayscale
            return np.dot(img[..., :3], [0.2989, 0.5870, 0.1140])
        return img


# Example usage
if __name__ == "__main__":
    # Simulate drift in microscopy images
    print("Creating synthetic test data with drift...")
    
    # Create a base image with features
    base_img = np.zeros((512, 512))
    y, x = np.ogrid[:512, :512]
    
    # Add some circular features
    center1 = (256, 256)
    center2 = (180, 350)
    center3 = (320, 200)
    
    mask1 = (x - center1[1])**2 + (y - center1[0])**2 < 50**2
    mask2 = (x - center2[1])**2 + (y - center2[0])**2 < 30**2
    mask3 = (x - center3[1])**2 + (y - center3[0])**2 < 40**2
    
    base_img[mask1] = 200
    base_img[mask2] = 150
    base_img[mask3] = 180
    
    # Add some noise
    base_img += np.random.normal(0, 10, base_img.shape)
    
    # Create sequence with random drift
    n_frames = 10
    images = []
    true_shifts = []
    
    for i in range(n_frames):
        # Simulate drift (cumulative)
        drift_y = np.cumsum(np.random.randn(i+1) * 2)[-1] if i > 0 else 0
        drift_x = np.cumsum(np.random.randn(i+1) * 2)[-1] if i > 0 else 0
        
        shifted = ndimage.shift(base_img, [drift_y, drift_x], order=1, mode='constant')
        shifted += np.random.normal(0, 5, shifted.shape)  # Add frame noise
        
        images.append(shifted)
        true_shifts.append([drift_y, drift_x])
    
    print(f"\nTrue cumulative shifts:\n{np.array(true_shifts)}\n")
    
    # Process images
    aligner = MicroscopyAligner(images)
    cropped = aligner.process_all(padding=10)
    
    print(f"\nDetected shifts:\n{aligner.shifts}\n")
    
    # Visualize
    aligner.visualize_results()