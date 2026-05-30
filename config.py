import torch
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

DATA_ROOT      = str(BASE_DIR / "archive") + "/"
TRAIN_CSV      = DATA_ROOT + "train.csv"
VAL_CSV        = DATA_ROOT + "val.csv"
TEST_CSV       = DATA_ROOT + "test.csv"
CHECKPOINT_DIR = str(BASE_DIR / "checkpoints") + "/"
LOG_DIR        = str(BASE_DIR / "logs") + "/"

IMG_HEIGHT   = 64
IMG_WIDTH    = 512
IMG_CHANNELS = 1

ALPHABET = (
    " !()*+,-./:;=?№"
    "0123456789"
    "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя"
)
BLANK_IDX   = 0
NUM_CLASSES = len(ALPHABET) + 1

CNN_CHANNELS     = [1, 64, 128, 256, 256, 512, 512]
CNN_KERNELS      = [3, 3, 3, 3, 3, 3]
CNN_STRIDES      = [1, 1, 1, 1, 1, 1]
POOL_CONFIGS     = [(2,2), (2,1), (2,1)]

RNN_HIDDEN       = 256
RNN_LAYERS       = 3
RNN_DROPOUT      = 0.2

ATTENTION_HEADS  = 8
ATTENTION_DIM    = 512

DEVICE        = "cuda" if torch.cuda.is_available() else "cpu"
BATCH_SIZE    = 32
NUM_EPOCHS    = 100
LR            = 3e-4
LR_SCHEDULER  = "cosine"
WARMUP_EPOCHS = 5
WEIGHT_DECAY  = 1e-4
GRAD_CLIP     = 5.0
PATIENCE      = 15
SEED          = 42

AUG_PROB         = 0.7
ELASTIC_ALPHA    = 50
ELASTIC_SIGMA    = 5
ROTATE_LIMIT     = 10
BLUR_LIMIT       = 3
NOISE_VAR        = 0.01
BRIGHTNESS_LIMIT = 0.3
CONTRAST_LIMIT   = 0.3
SHEAR_RANGE      = 10

BEAM_WIDTH = 10
LM_WEIGHT  = 0.3
BEST_MODEL = CHECKPOINT_DIR + "best_model.pth"
