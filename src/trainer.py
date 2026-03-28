"""Generic PyTorch training loop with early stopping, LR scheduling, and output saving."""

import json
import numpy as np
import torch
import torch.nn as nn
import torch.optim as optim
from pathlib import Path

from .config import TrainConfig, PathConfig


def train_model(
    model: nn.Module,
    train_loader,
    val_loader,
    train_cfg: TrainConfig,
    path_cfg: PathConfig,
    model_name: str,
    device: torch.device,
) -> dict:
    """Train a PyTorch model and save checkpoints + loss history.

    Returns:
        dict with 'train_losses', 'val_losses'
    """
    criterion = nn.MSELoss()
    optimizer = optim.Adam(model.parameters(), lr=train_cfg.learning_rate)
    scheduler = optim.lr_scheduler.ReduceLROnPlateau(
        optimizer, mode="min", factor=train_cfg.lr_factor, patience=train_cfg.lr_patience
    )

    best_val_loss = float("inf")
    best_state = None
    patience_counter = 0
    train_losses, val_losses = [], []

    print(f"\n{'='*60}")
    print(f"  Training: {model_name}")
    print(f"  Params: {sum(p.numel() for p in model.parameters()):,}")
    print(f"{'='*60}")

    for epoch in range(train_cfg.epochs):
        # --- Train ---
        model.train()
        epoch_loss, n_batches = 0.0, 0
        for X_batch, Y_batch in train_loader:
            optimizer.zero_grad()
            preds = model(X_batch)
            loss = criterion(preds, Y_batch)
            loss.backward()
            torch.nn.utils.clip_grad_norm_(model.parameters(), max_norm=train_cfg.grad_clip)
            optimizer.step()
            epoch_loss += loss.item()
            n_batches += 1
        train_loss = epoch_loss / n_batches
        train_losses.append(train_loss)

        # --- Validate ---
        model.eval()
        val_loss, n_val = 0.0, 0
        with torch.no_grad():
            for X_batch, Y_batch in val_loader:
                preds = model(X_batch)
                val_loss += criterion(preds, Y_batch).item()
                n_val += 1
        val_loss /= n_val
        val_losses.append(val_loss)

        scheduler.step(val_loss)
        lr = optimizer.param_groups[0]["lr"]

        if (epoch + 1) % 10 == 0 or epoch == 0:
            print(f"  Epoch {epoch+1:3d}/{train_cfg.epochs} | "
                  f"Train: {train_loss:.6f} | Val: {val_loss:.6f} | LR: {lr:.6f}")

        # Early stopping
        if val_loss < best_val_loss:
            best_val_loss = val_loss
            best_state = {k: v.cpu().clone() for k, v in model.state_dict().items()}
            patience_counter = 0
        else:
            patience_counter += 1
            if patience_counter >= train_cfg.patience:
                print(f"  Early stopping at epoch {epoch+1}")
                break

    # Restore best weights
    model.load_state_dict(best_state)
    model.to(device)
    print(f"  Best val loss: {best_val_loss:.6f}")

    # --- Save model ---
    model_path = path_cfg.model_dir(model_name) / f"{model_name}.pt"
    torch.save(model.state_dict(), model_path)
    print(f"  Model saved: {model_path}")

    # --- Save loss history ---
    loss_data = {"train_loss": train_losses, "val_loss": val_losses}
    loss_path = path_cfg.loss_dir(model_name) / "loss_history.json"
    with open(loss_path, "w") as f:
        json.dump(loss_data, f, indent=2)

    # Save loss as CSV too (convenient for later analysis)
    import pandas as pd
    loss_df = pd.DataFrame({"epoch": range(1, len(train_losses) + 1),
                            "train_loss": train_losses, "val_loss": val_losses})
    loss_df.to_csv(path_cfg.loss_dir(model_name) / "loss_history.csv", index=False)

    return loss_data


def predict(model: nn.Module, X_tensor: torch.Tensor) -> np.ndarray:
    """Run inference, return numpy array."""
    model.eval()
    with torch.no_grad():
        preds = model(X_tensor).cpu().numpy()
    return preds
