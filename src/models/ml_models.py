"""Scikit-learn based ML models: Isolation Forest, Random Forest."""

import numpy as np
from sklearn.ensemble import IsolationForest, RandomForestRegressor

from ..config import ModelConfig


def train_isolation_forest(train_df, cfg: ModelConfig):
    model = IsolationForest(
        n_estimators=cfg.iso_n_estimators,
        contamination=cfg.iso_contamination,
        max_samples="auto",
        random_state=42,
        n_jobs=-1,
    )
    model.fit(train_df)
    return model


def predict_isolation_forest(model, test_df):
    preds = model.predict(test_df)
    scores = model.decision_function(test_df)
    anomalies = (preds == -1).astype(int)
    return anomalies, scores


def train_random_forest(X_train, y_train, cfg: ModelConfig):
    model = RandomForestRegressor(
        n_estimators=cfg.rf_n_estimators,
        max_depth=cfg.rf_max_depth,
        min_samples_split=5,
        random_state=42,
        n_jobs=-1,
    )
    model.fit(X_train, y_train)
    return model
