from checkpoint import save_checkpoint, load_checkpoint, clear_checkpoint

save_checkpoint(
    phase="phase1",
    hospital_index=25,
    hospital_name="Apollo Hospital"
)

data = load_checkpoint()

print(data)

clear_checkpoint()