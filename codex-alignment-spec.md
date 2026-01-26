# Microscopy Image Drift Correction - Technical Specification

## Objective
Correct translational stage drift in time-lapse microscopy sequences and crop to maximum common valid region.

## Input/Output
**Input:** N grayscale/RGB images with 2D translational drift (Δx, Δy)  
**Output:** N aligned images, cropped to identical dimensions, maximum retained area

## Core Algorithm: Enhanced Correlation Coefficient (ECC)

### Why ECC?
- **Subpixel accuracy** (0.01-0.1 pixel precision)
- **Robust to intensity variations** (handles photobleaching, exposure changes)
- **Optimal speed/accuracy tradeoff** for microscopy
- **Proven method** in computer vision (OpenCV implementation)

### Mathematical Basis
Maximizes normalized cross-correlation between reference and target:

```
ECC(Δx, Δy) = Σ[I_ref · I_target(shifted)] / √[Σ I_ref² · Σ I_target²]
```

Solved via iterative Gauss-Newton optimization.

## Implementation

### 1. Compute Shifts (ECC Method)
```python
import cv2
import numpy as np

def compute_shifts(images, reference_idx=0):
    """Returns (N, 2) array of [dy, dx] shifts"""
    reference = to_grayscale_float32(images[reference_idx])
    shifts = np.zeros((len(images), 2))
    
    criteria = (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, 100, 1e-6)
    
    for i, img in enumerate(images):
        if i == reference_idx:
            continue
        
        target = to_grayscale_float32(img)
        warp_matrix = np.eye(2, 3, dtype=np.float32)
        
        _, warp_matrix = cv2.findTransformECC(
            reference, target, warp_matrix,
            motionType=cv2.MOTION_TRANSLATION,
            criteria=criteria
        )
        
        shifts[i] = [warp_matrix[1, 2], warp_matrix[0, 2]]  # [dy, dx]
    
    return shifts
```

### 2. Align Images
```python
from scipy import ndimage

def align_images(images, shifts, order=3):
    """Apply inverse shifts using scipy interpolation"""
    aligned = []
    for img, shift in zip(images, shifts):
        # Negative shift aligns TO reference
        if img.ndim == 2:
            aligned.append(ndimage.shift(img, -shift, order=order, cval=0))
        else:
            # RGB: shift each channel
            result = np.zeros_like(img)
            for c in range(img.shape[2]):
                result[:, :, c] = ndimage.shift(img[:, :, c], -shift, order=order, cval=0)
            aligned.append(result)
    return aligned
```

### 3. Compute Crop Bounds
```python
def compute_crop_bounds(shifts, image_shape, padding=5):
    """Calculate maximum rectangle containing valid data in all frames"""
    H, W = image_shape[:2]
    
    # Shift extremes define crop region
    top = int(np.ceil(max(0, np.max(shifts[:, 0])))) + padding
    bottom = H - int(np.ceil(max(0, -np.min(shifts[:, 0])))) - padding
    left = int(np.ceil(max(0, np.max(shifts[:, 1])))) + padding
    right = W - int(np.ceil(max(0, -np.min(shifts[:, 1])))) - padding
    
    # Ensure valid bounds
    top = max(0, min(top, H))
    bottom = max(top + 1, min(bottom, H))
    left = max(0, min(left, W))
    right = max(left + 1, min(right, W))
    
    return (top, bottom, left, right)
```

### 4. Crop All Images
```python
def crop_images(images, crop_bounds):
    """Crop to common region"""
    top, bottom, left, right = crop_bounds
    return [img[top:bottom, left:right] if img.ndim == 2 
            else img[top:bottom, left:right, :] 
            for img in images]
```

## Complete Pipeline
```python
def process_sequence(images, reference_idx=0, padding=5, interpolation_order=3):
    # 1. Compute translational shifts
    shifts = compute_shifts(images, reference_idx)
    
    # 2. Align images
    aligned = align_images(images, shifts, order=interpolation_order)
    
    # 3. Compute crop bounds
    bounds = compute_crop_bounds(shifts, images[0].shape, padding)
    
    # 4. Crop to common area
    cropped = crop_images(aligned, bounds)
    
    return cropped, shifts, bounds
```

## Preprocessing Requirements
```python
def to_grayscale_float32(img):
    """Convert image to format required by ECC"""
    if img.ndim == 3:
        gray = np.dot(img[..., :3], [0.2989, 0.5870, 0.1140])
    else:
        gray = img
    return gray.astype(np.float32)
```

## Error Handling
```python
try:
    _, warp_matrix = cv2.findTransformECC(...)
except cv2.error:
    # Fallback: phase correlation (faster, less robust)
    from skimage.registration import phase_cross_correlation
    shift, _, _ = phase_cross_correlation(reference, target, upsample_factor=10)
```

## Alternative Methods (If ECC Fails)

| Method | Use Case | Speed | Robustness |
|--------|----------|-------|------------|
| **ECC** | Default choice | Medium | High |
| Phase Correlation | Clean data, speed critical | Fast | Medium |
| Template Matching | Noisy/sparse features | Slow | Highest |

### Phase Correlation (Fallback)
```python
from skimage.registration import phase_cross_correlation

shift, error, _ = phase_cross_correlation(
    reference, target, upsample_factor=10
)
```

### Template Matching (Maximum Robustness)
```python
from skimage.feature import match_template

# Use central 60% as template
margin = int(0.2 * min(H, W))
template = reference[margin:-margin, margin:-margin]
result = match_template(target, template, pad_input=True)
ij = np.unravel_index(np.argmax(result), result.shape)
shift = [ij[0] - H//2, ij[1] - W//2]
```

## Key Parameters

| Parameter | Recommended | Description |
|-----------|-------------|-------------|
| `reference_idx` | 0 or N//2 | Reference frame (middle frame often best) |
| `padding` | 5-10 pixels | Safety margin for cropping |
| `interpolation_order` | 3 | Cubic interpolation (0=nearest, 1=linear) |
| `max_iterations` | 100 | ECC convergence iterations |
| `termination_eps` | 1e-6 | ECC convergence threshold |

## Performance Characteristics

**Speed:** ~0.5-2 seconds per frame pair (512×512 images, modern CPU)  
**Accuracy:** Subpixel (typically 0.01-0.1 pixel)  
**Typical crop loss:** 5-15% of area (depends on drift magnitude)  
**Memory:** O(N × H × W) where N = number of frames

## Validation

Check registration quality:
```python
def validate_registration(reference, aligned):
    ref_norm = (reference - reference.mean()) / reference.std()
    align_norm = (aligned - aligned.mean()) / aligned.std()
    ncc = np.mean(ref_norm * align_norm)
    return ncc  # Should be > 0.90 for good registration
```

## Dependencies
- `numpy` - Array operations
- `scipy` - Image interpolation (ndimage.shift)
- `opencv-python` (cv2) - ECC algorithm
- `scikit-image` - Phase correlation, template matching (fallback methods)

## Critical Implementation Notes

1. **Image format:** ECC requires float32 grayscale images
2. **Shift direction:** Apply **negative** shift to align images TO reference
3. **Interpolation:** Use order=3 (cubic) for best quality, order=1 for speed
4. **Crop bounds:** Add padding to avoid edge artifacts from interpolation
5. **Reference frame:** Middle frame often better than first (less cumulative drift)

## Output Guarantees

- All output images have identical dimensions (H_crop × W_crop)
- All pixels in output contain valid interpolated data (no black borders)
- Crop region is the **maximum** rectangle valid across all frames
- Alignment accuracy is subpixel (limited by interpolation order)