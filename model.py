import torch
import torch.nn as nn
import torch.nn.functional as F
from config import *


class GatedConvBlock(nn.Module):
    def __init__(self, in_ch: int, out_ch: int,
                 kernel: int = 3, stride: int = 1,
                 pool: tuple | None = None,
                 dropout: float = 0.1):
        super().__init__()
        if isinstance(kernel, tuple):
            pad = tuple(k // 2 for k in kernel)
        else:
            pad = kernel // 2
        self.conv_main = nn.Conv2d(in_ch, out_ch, kernel, stride=stride,
                                   padding=pad, bias=False)
        self.bn_main = nn.BatchNorm2d(out_ch)
        self.act = nn.PReLU(out_ch)
        self.conv_gate = nn.Conv2d(in_ch, out_ch, kernel, stride=stride,
                                   padding=pad, bias=False)
        self.bn_gate = nn.BatchNorm2d(out_ch)
        self.pool = nn.MaxPool2d(pool) if pool else None
        self.drop = nn.Dropout2d(dropout)

    def forward(self, x):
        main = self.act(self.bn_main(self.conv_main(x)))
        gate = torch.sigmoid(self.bn_gate(self.conv_gate(x)))
        out = main * gate
        if self.pool:
            out = self.pool(out)
        return self.drop(out)


class ResidualGatedBlock(nn.Module):
    def __init__(self, ch: int, dropout: float = 0.1):
        super().__init__()
        self.gated = GatedConvBlock(ch, ch, dropout=dropout)
        self.norm = nn.BatchNorm2d(ch)

    def forward(self, x):
        return self.norm(x + self.gated(x))


class FeatureExtractor(nn.Module):
    def __init__(self):
        super().__init__()
        self.block1 = GatedConvBlock(1, 64, 3, pool=(2, 2), dropout=0.1)
        self.res1 = ResidualGatedBlock(64, dropout=0.1)
        self.block2 = GatedConvBlock(64, 128, 3, pool=(2, 1), dropout=0.1)
        self.res2 = ResidualGatedBlock(128, dropout=0.1)
        self.block3 = GatedConvBlock(128, 256, 3, pool=(2, 1), dropout=0.15)
        self.res3 = ResidualGatedBlock(256, dropout=0.15)
        self.block4 = GatedConvBlock(256, 256, (2, 4), dropout=0.15)
        self.block5 = GatedConvBlock(256, 512, 3, pool=(2, 1), dropout=0.2)
        self.res4 = ResidualGatedBlock(512, dropout=0.2)
        self.block6 = GatedConvBlock(512, 512, (2, 4), dropout=0.2)

    def forward(self, x):
        x = self.block1(x)
        x = self.res1(x)
        x = self.block2(x)
        x = self.res2(x)
        x = self.block3(x)
        x = self.res3(x)
        x = self.block4(x)
        x = self.block5(x)
        x = self.res4(x)
        x = self.block6(x)
        b, c, h, w = x.shape
        x = x.permute(0, 3, 1, 2)
        x = x.contiguous().view(b, w, c * h)
        return x


class PositionalEncoding(nn.Module):
    def __init__(self, d_model: int, max_len: int = 1024, dropout: float = 0.1):
        super().__init__()
        self.dropout = nn.Dropout(dropout)
        pe = torch.zeros(max_len, d_model)
        pos = torch.arange(max_len).unsqueeze(1)
        div = torch.exp(
            torch.arange(0, d_model, 2) * (-torch.log(torch.tensor(10000.0)) / d_model)
        )
        pe[:, 0::2] = torch.sin(pos * div)
        pe[:, 1::2] = torch.cos(pos * div)
        self.register_buffer("pe", pe.unsqueeze(0))

    def forward(self, x):
        return self.dropout(x + self.pe[:, :x.size(1)])


class BiGRUEncoder(nn.Module):
    def __init__(self, input_size: int):
        super().__init__()
        self.gru = nn.GRU(
            input_size=input_size,
            hidden_size=RNN_HIDDEN,
            num_layers=RNN_LAYERS,
            batch_first=True,
            bidirectional=True,
            dropout=RNN_DROPOUT if RNN_LAYERS > 1 else 0.0,
        )
        self.norm = nn.LayerNorm(RNN_HIDDEN * 2)

    def forward(self, x):
        out, _ = self.gru(x)
        return self.norm(out)


class AttentionRefinement(nn.Module):
    def __init__(self, d_model: int, nhead: int = 8, ff_dim: int = 1024):
        super().__init__()
        self.pos_enc = PositionalEncoding(d_model)
        self.self_attn = nn.MultiheadAttention(d_model, nhead, dropout=0.1,
                                               batch_first=True)
        self.ffn = nn.Sequential(
            nn.Linear(d_model, ff_dim),
            nn.GELU(),
            nn.Dropout(0.1),
            nn.Linear(ff_dim, d_model),
        )
        self.norm1 = nn.LayerNorm(d_model)
        self.norm2 = nn.LayerNorm(d_model)
        self.drop = nn.Dropout(0.1)

    def forward(self, x):
        x = self.pos_enc(x)
        attn_out, _ = self.self_attn(x, x, x)
        x = self.norm1(x + self.drop(attn_out))
        x = self.norm2(x + self.drop(self.ffn(x)))
        return x


class RussianHTR(nn.Module):
    def __init__(self):
        super().__init__()
        self.cnn = FeatureExtractor()
        with torch.no_grad():
            dummy = torch.zeros(1, IMG_CHANNELS, IMG_HEIGHT, IMG_WIDTH)
            cnn_out = self.cnn(dummy)
            cnn_feat_size = cnn_out.shape[-1]

        self.proj = nn.Sequential(
            nn.Linear(cnn_feat_size, ATTENTION_DIM),
            nn.LayerNorm(ATTENTION_DIM),
            nn.Dropout(0.2),
        )
        self.bigru = BiGRUEncoder(ATTENTION_DIM)
        self.gru_proj = nn.Linear(RNN_HIDDEN * 2, ATTENTION_DIM)
        self.attention = AttentionRefinement(ATTENTION_DIM, nhead=ATTENTION_HEADS)
        self.classifier = nn.Linear(ATTENTION_DIM, NUM_CLASSES)

    def forward(self, x):
        feat = self.cnn(x)
        feat = self.proj(feat)
        feat = self.bigru(feat)
        feat = self.gru_proj(feat)
        feat = self.attention(feat)
        logits = self.classifier(feat)
        return logits.permute(1, 0, 2)


def build_model() -> RussianHTR:
    model = RussianHTR().to(DEVICE)
    total = sum(p.numel() for p in model.parameters())
    print(f"[Model] Параметров всего: {total:,}")
    print(f"[Model] Устройство: {DEVICE}")
    return model
