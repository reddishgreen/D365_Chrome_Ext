// Simple icon generator using base64 encoded PNGs
const fs = require('fs');
const path = require('path');

// 16x16 icon - Blue square with yellow lightning
const icon16 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAC/SURBVDiN7ZMxDsIwDEVfUgZWJC7ABbgMl+EycAE4AisTKhIL/w4JQ1WqCrHAJ1lK7Pj9xHYE/CkBTICpMZ4Bc2AE9B7AClgDm2b8CGyBHdADFsCxAWwIvxNwAI4VQLQFrgQvgHsFEG2BK8ELYF8BRFvgSvACWFUAH5rTSvDc3X0I7+LO7uDm7m4VQOkuKsGc7tYBxGbgSvAOmFUAYjNwJXgHzCsAsRm4ErwDFhXAC+P32gPfAU/gs/Z+018qfQEk003Jo3lQmwAAAABJRU5ErkJggg==',
  'base64'
);

// 48x48 icon - Blue square with yellow lightning
const icon48 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAADdgAAA3YBfdWCzAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAIgSURBVGiB7Zk9aBRBFMd/s3t7l1yMIKhgY2EhphALsbAQwcJCEMHCQhDBwkIEC0GwsLAQBAsLC0GwsBAECwsRLCxEsLAQBAsLESws5OzuzvPYu2Qvm929m929S/YHC7Mz82bm/96bmZ0dWGeddcD/RqoGzAJTwDgwBowCo8BoOY4CIy3tRoCRtliTOgDmgfkyN2/mOAfMAbOBgQWC+7q4oZbOiNW8O3Hf1nyZmwNmAtf+EZhsOZ+szpukRsz5pE1f0zjZcj5Zjy+AiZbzCaARuHZLoNGqcLJqQoVVzW8VxKoElr8KYhUC/v1VEMsWCPX/qyDWLdBN/18FsUqBXvRXQaxKoFf9VRCrEOhVfxXEqgX60V8FsUqBfvVXQaxKYBD9VRCrEBhUfxVE+ToAdeCUi0+5XAD6qr/XteO0TjWmz7BQBxrAZeAAMObiB+C+i+d9bFhDfD+Wd0x/EF4Cuy4+BewCL6I6JoBbQd0jBDg+tNBRF/sCL27Bsy+p62u8fJIWP9QW+OX684nu2B7gWOD6MnBF4BJw2cevhOq+BIaB0x25I2WFauAwcD6oewgYLnO+70chBPZ0KzsEHOzIHQzi+TZdP+I9d3EfPw7sD+r2Y1c82u6AexrYAWaAbZf7BGy5+BZQD9S1DWy6vM/t+J44b9eqK/AduBPU9R24C3wDvgIPgAdl7hF46PIBfAXuBa79AShyt93Qs585AAAAAElFTkSuQmCC',
  'base64'
);

// 128x128 icon - Blue square with yellow lightning
const icon128 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAIAAAACACAYAAADDPmHLAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAADsQAAA7EB9YPtSQAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAWRSURBVHic7Z09aBRBFMd/c3t3uUtMYqGCgoWFhZhCLMTCQgQLCxEsLAQRLCxEsLAQBAsLC0GwsLAQBAsLQbCwEAQLCxEsLETwQxDvdneexa3J5e52d2/3bjO7+/7gcMzOvHnz/29m3rydXYCoqKioqCjLiJwB1MJ+4CAwAhwAhoEhYAgYBIaAQWCgbTPQth1sa96Owm+AWeAn8KP03gVmgGlgqoynWjPTsxlo2gWjEieBs8Bx4CjQ59gfwFfgi/RfgHHgHfAR+Ky1XnbsV0IE3APOAWekdxr9BN5K/wF4C3zo5TBLIoDLwBXgNDDq2D4D76V/I/2s1npfRcMmqCWQAueB68B5x/YFeCX9C+nfaK1rFe2bpJaAfuAacBW44Ng+AC+kfwnGA9ZbrStV7VugloD9wE3gLnDcsU0BT6V/Jr20FytW/QBq+P8S0O0LPw1cAy4DB5xxk8BT6R8Dk6HBNkjKCeBGV1o/cAu4I/0Ox/ZR+kdSK7z/GwPKEngjEz6L8O8Ad4ER5+kJ4JH0j4BPseG2iNQRwJVOfX8YuCf9PuC2MzYOPJD+PpDLgNdqFkkngeHSnAY+kI/yLgL3pR8Fbjtj3wAPpb8PfG8xzkZJdUW/AjwAPoX+X8DPkv8d+Lsl2QnQ+A/gSqe+/whwQ/peZ+wl8Fj6+8Dn0DAb3+LbcZb8B3APeBo6fhpYkr4IPG8p1vqLnSAjkg0pV+XY8dJ2DHjh2L4Bb0rjxkvbkWP/XNqed2xHpI+JvhW8s0R+BF5Ir7WuxYRfB1YcfzAU2xXg1aY7TgCvpY+ZANz1wQz5evAVECTQJRB+Avg4yRNAKPUE4P8b+FvkrwH3S/4WztiSY1sEnsm/jz+t/JV/vnR/UfpZx7bsjH0ix5p25vPXgXmpBem/S+1I/9WxLUovyH/fODbfthUrgM8SD+RYwfFvwrfEfSnVltyPx7ak1qV//LdtA78APyX2pXbjX8e3SVkBfJRYBX45tiWp1RL/OvZNWAFMSq04fZdK/mvytSO1Ij3k63ta+nL8C/KvYd8Gse8B7El8ljohtS/nX8e+iT3H3i99T+pI6kDqUOpQ6kj6fdLH0gv4S8V65JdP5RXDJekFJ/4F6Rcc22HHvhy/fzX+I8e+D+eJZU/qSOpo05jUoeMfk16Afx3/cdkNfJbT/ivgDPDMGfsNWHT8Z4C/0p8Enkg/D5wCTkv/G3gOPJX+L/kE8kz638AZ8l7kV+mfkz9mvgRU/pz8TvFf5G/tFJz+Fce/SP6izjj+Qr528i/1BG3AX+IfN2ECqOw/xnLnr0uGJdYdWzH+ddLVAAoBXJJYA+adsSUnvgV4g1QX/gL+y/rz5PPs+X/Jr+e7KyL8+VMAtW4C2HD6Lu2+/Ev4x/8DfgN/nf5fSvV9m9j2CaC4JVtw7rk7/gL+G7kryX/7GkEtAU3y/fx1/OfBFPmx38R/Qv+S2Pb9bJf9gI5m8Jf/d//E/wb/w5+Of53aAqowBFz6H/7vSm37S8F/ie0vBf8ltv2lYDUCuAKcq9CfI/9pKESgfHtwI/QH/L/XA/lP/SH2q/zNgr+AP7vXqgVQdQmYBJ6Rfw+vEn8V/8OLX/x16glwdwdD+iv6n/gf9EL4y+xKAAXqCtDkO/4V/Hf8IQLdUFsAu0jtT8I+cQL4G/3E8R9tE/sWaQfwY/1/y/439DFRRwCh3/u3k/6W7v1bxe8lYraCQ7/v0Qr+r327xb8TJy4BQwH+28W/UwI6gu2yBHSLv+gE4DsBuHMCcJ0AXCcA1wnAdQJwnQBcJwB3xwng7wQQ+r1/u/m3qz9jt4IhAmO/97ed/Nv13w7/uKioqKioqFbpH7aM7xJvi5klAAAAAElFTkSuQmCC',
  'base64'
);

// Write the icons
fs.writeFileSync(path.join(__dirname, 'icons', 'icon16.png'), icon16);
fs.writeFileSync(path.join(__dirname, 'icons', 'icon48.png'), icon48);
fs.writeFileSync(path.join(__dirname, 'icons', 'icon128.png'), icon128);

console.log('Icons created successfully!');
console.log('- icons/icon16.png');
console.log('- icons/icon48.png');
console.log('- icons/icon128.png');
