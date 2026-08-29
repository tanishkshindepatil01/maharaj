from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np


OUTPUT_DIR = Path(__file__).resolve().parents[1] / "figures"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

NAVY = "#2F5D8A"
BLUE = "#5B8DB8"
ORANGE = "#F28E2B"
GREEN = "#59A14F"
TEAL = "#4EAAA8"
GREY = "#6B7280"
GRID = "#D9DEE5"


def style_axis(axis, grid_axis="y"):
    axis.spines["top"].set_visible(False)
    axis.spines["right"].set_visible(False)
    axis.spines["left"].set_color("#AEB7C2")
    axis.spines["bottom"].set_color("#AEB7C2")
    axis.tick_params(colors="#4B5563", labelsize=9)
    axis.grid(axis=grid_axis, color=GRID, linewidth=0.8, alpha=0.75)
    axis.set_axisbelow(True)


def add_vertical_labels(axis, bars, fmt):
    for bar in bars:
        value = bar.get_height()
        axis.text(
            bar.get_x() + bar.get_width() / 2,
            value,
            fmt.format(value),
            ha="center",
            va="bottom",
            fontsize=9,
            color="#1F2937",
        )


def add_horizontal_labels(axis, bars, fmt, pad_fraction=0.018):
    upper = axis.get_xlim()[1]
    for bar in bars:
        value = bar.get_width()
        axis.text(
            value + upper * pad_fraction,
            bar.get_y() + bar.get_height() / 2,
            fmt.format(value),
            ha="left",
            va="center",
            fontsize=8.5,
            color="#1F2937",
        )


def create_model_efficiency_charts():
    models = ["YOLOv8n", "YOLO11n", "YOLO11s"]
    model_size_mb = [6.00, 5.21, 18.35]
    median_fps = [8.53, 8.35, 8.34]
    validation_map = [0.4288, 0.4002, 0.4107]
    colors = [ORANGE, NAVY, BLUE]
    x = np.arange(len(models))

    specifications = [
        (
            "figure_07_model_size.png",
            "Stored model size",
            "Model size (MB)",
            model_size_mb,
            20,
            "{:.2f}",
        ),
        (
            "figure_07_inference_speed.png",
            "Validation inference speed",
            "Median end-to-end FPS",
            median_fps,
            9.3,
            "{:.2f}",
        ),
        (
            "figure_07_validation_map.png",
            "Validation localisation quality",
            "mAP50-95",
            validation_map,
            0.46,
            "{:.4f}",
        ),
    ]

    for filename, title, ylabel, values, ymax, label_format in specifications:
        fig, axis = plt.subplots(figsize=(4.75, 3.35))
        bars = axis.bar(x, values, color=colors, width=0.64)
        axis.set_title(title, loc="left", fontsize=11.5, weight="bold")
        axis.set_ylabel(ylabel)
        axis.set_ylim(0, ymax)
        axis.set_xticks(x, models)
        add_vertical_labels(axis, bars, label_format)
        style_axis(axis)
        fig.tight_layout(pad=0.8)
        fig.savefig(OUTPUT_DIR / filename, dpi=300, facecolor="white", bbox_inches="tight")
        plt.close(fig)


def create_table_4_1_charts():
    representations = [
        "PyTorch FP32",
        "TFLite FP32",
        "TFLite INT8",
        "NCNN FP32",
        "NCNN FP16",
    ]
    colors = [NAVY, BLUE, ORANGE, GREEN, TEAL]
    latency = [397.48, 143.70, 44.13, 75.09, 72.31]
    fps = [2.52, 6.95, 22.62, 13.26, 13.82]
    cpu = [44.47, 69.95, 62.83, 83.32, 87.08]
    ram = [839.47, 868.27, 824.22, 835.25, 835.45]
    precision = [0.716, 0.669, 0.661, 0.667, 0.672]
    recall = [0.550, 0.542, 0.539, 0.542, 0.542]
    f1 = [0.622, 0.599, 0.594, 0.598, 0.600]
    map50 = [0.548, 0.535, 0.503, 0.535, 0.537]
    map5095 = [0.275, 0.270, 0.234, 0.271, 0.270]

    y = np.arange(len(representations))

    def horizontal_panel(axis, values, title, xlabel, xmax, label_format):
        bars = axis.barh(y, values, color=colors, height=0.62)
        axis.set_title(title, loc="left", fontsize=11, weight="bold")
        axis.set_xlabel(xlabel)
        axis.set_xlim(0, xmax)
        axis.set_yticks(y, representations)
        axis.invert_yaxis()
        add_horizontal_labels(axis, bars, label_format)
        style_axis(axis, grid_axis="x")

    fig, axes = plt.subplots(1, 2, figsize=(13.6, 3.55))
    horizontal_panel(axes[0], latency, "Mean inference latency", "Milliseconds per image", 430, "{:.1f}")
    horizontal_panel(axes[1], fps, "Median throughput", "Frames per second", 25, "{:.2f}")
    fig.tight_layout(pad=0.8, w_pad=2.5)
    fig.savefig(
        OUTPUT_DIR / "figure_table_4_1_runtime.png",
        dpi=300,
        facecolor="white",
        bbox_inches="tight",
    )
    plt.close(fig)

    fig, axes = plt.subplots(1, 2, figsize=(13.6, 3.55))
    horizontal_panel(axes[0], cpu, "Mean CPU usage", "CPU usage (%)", 100, "{:.1f}")
    horizontal_panel(axes[1], ram, "Peak process memory", "Peak RAM (MB)", 930, "{:.1f}")
    fig.tight_layout(pad=0.8, w_pad=2.5)
    fig.savefig(
        OUTPUT_DIR / "figure_table_4_1_resources.png",
        dpi=300,
        facecolor="white",
        bbox_inches="tight",
    )
    plt.close(fig)

    fig, axes = plt.subplots(1, 2, figsize=(13.6, 3.85))
    x = np.arange(len(representations))
    width = 0.23
    metric_axis = axes[0]
    metric_axis.bar(x - width, precision, width, label="Precision", color=NAVY)
    metric_axis.bar(x, recall, width, label="Recall", color=ORANGE)
    metric_axis.bar(x + width, f1, width, label="F1", color=GREEN)
    metric_axis.set_title("Operating-point detection metrics", loc="left", fontsize=11, weight="bold")
    metric_axis.set_ylabel("Score")
    metric_axis.set_ylim(0, 0.8)
    metric_axis.set_xticks(x, representations, rotation=14, ha="right")
    metric_axis.legend(frameon=False, fontsize=8, ncol=3, loc="upper right")
    style_axis(metric_axis)

    map_axis = axes[1]
    width = 0.34
    map_axis.bar(x - width / 2, map50, width, label="mAP50", color=NAVY)
    map_axis.bar(x + width / 2, map5095, width, label="mAP50-95", color=TEAL)
    map_axis.set_title("Mean average precision", loc="left", fontsize=11, weight="bold")
    map_axis.set_ylabel("mAP score")
    map_axis.set_ylim(0, 0.62)
    map_axis.set_xticks(x, representations, rotation=14, ha="right")
    map_axis.legend(frameon=False, fontsize=8, ncol=2, loc="upper right")
    style_axis(map_axis)
    fig.tight_layout(pad=0.8, w_pad=2.5)
    fig.savefig(
        OUTPUT_DIR / "figure_table_4_1_accuracy.png",
        dpi=300,
        facecolor="white",
        bbox_inches="tight",
    )
    plt.close(fig)


def create_deployment_tradeoff_scatter():
    representations = [
        "PyTorch FP32",
        "TFLite FP32",
        "TFLite INT8",
        "NCNN FP32",
        "NCNN FP16",
    ]
    fps = [2.52, 6.95, 22.62, 13.26, 13.82]
    map5095 = [0.275, 0.270, 0.234, 0.271, 0.270]
    colors = [NAVY, BLUE, ORANGE, GREEN, TEAL]

    fig, axis = plt.subplots(figsize=(7.4, 4.25))
    axis.scatter(
        fps,
        map5095,
        s=125,
        c=colors,
        edgecolors="white",
        linewidths=1.2,
        zorder=3,
    )

    label_offsets = {
        "PyTorch FP32": (10, 8),
        "TFLite FP32": (-12, -20),
        "TFLite INT8": (-10, 10),
        "NCNN FP32": (-34, 13),
        "NCNN FP16": (18, -19),
    }
    alignments = {
        "TFLite FP32": "right",
        "TFLite INT8": "right",
        "NCNN FP32": "right",
    }

    for name, x_value, y_value, color in zip(representations, fps, map5095, colors):
        offset = label_offsets[name]
        axis.annotate(
            f"{name}\n{x_value:.2f} FPS, {y_value:.3f}",
            (x_value, y_value),
            xytext=offset,
            textcoords="offset points",
            ha=alignments.get(name, "left"),
            va="center",
            fontsize=8.5,
            color="#1F2937",
            arrowprops={"arrowstyle": "-", "color": color, "lw": 0.8},
        )

    axis.set_title("Deployment speed-accuracy trade-off", loc="left", fontsize=12, weight="bold")
    axis.set_xlabel("Median throughput (FPS)")
    axis.set_ylabel("mAP50-95")
    axis.set_xlim(0, 25)
    axis.set_ylim(0.225, 0.285)
    style_axis(axis, grid_axis="both")
    axis.annotate(
        "Faster and more accurate",
        xy=(24.0, 0.282),
        xytext=(16.3, 0.282),
        ha="left",
        va="center",
        fontsize=8.5,
        color=GREY,
        arrowprops={"arrowstyle": "->", "color": GREY, "lw": 1.0},
    )
    fig.tight_layout(pad=0.8)
    fig.savefig(
        OUTPUT_DIR / "figure_deployment_speed_accuracy_scatter.png",
        dpi=300,
        facecolor="white",
        bbox_inches="tight",
    )
    plt.close(fig)


if __name__ == "__main__":
    create_model_efficiency_charts()
    create_table_4_1_charts()
    create_deployment_tradeoff_scatter()
