#!/usr/bin/python
import argparse
from run_sagabz_socio import RunSagzab, RunSocio
import os

FORMAT_PATH = os.getcwd() + r"\Formats\socio_format_new.docx"

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--run-socio", action="store_true", default=False, help="run the socio option"
    )
    parser.add_argument(
        "--run-sagabz", action="store_true", default=False, help="run the sagabz option"
    )

    parser.add_argument(
        "--run-combine-data",
        action="store_true",
        default=False,
        help="when used, the data will be combined_data.xlsx",
    )

    parser.add_argument(
        "--run-commander-file",
        action="store_true",
        default=False,
        help="when used, commander file graphs will be generated.",
    )

    parser.add_argument(
        "--dont-run-classification",
        action="store_true",
        default=False,
        help="when used, ai classifications will not execute.",
    )

    parser.add_argument(
        "--socio-output-path",
        type=str,
        default=r"sociometry_output",
        help="path for socio ouptut",
    )
    parser.add_argument(
        "--raw-data-path", type=str, default=r"Excels", help="path for excels files."
    )
    parser.add_argument(
        "--names-to-hashes",
        type=str,
        default=r"False",
        help="convert names to hashes. for unanimous data.",
    )

    args = parser.parse_args()

    format_path = FORMAT_PATH

    relative_path = os.getcwd()
    socio_output_path = os.path.join(relative_path, args.socio_output_path)

    run_classification = not args.dont_run_classification

    raw_data_path = os.path.join(relative_path, args.raw_data_path)

    if args.run_socio:
        run_obj = RunSocio(
            raw_data_path,
            socio_output_path,
            format_path,
            testing=True,
            combine_excels=args.run_combine_data,
            run_classification=run_classification,
            names_to_hashes=args.names_to_hashes,
            is_commander_file=args.run_commander_file,
        )
        run_obj.run()

    elif args.run_sagabz:
        run_obj = RunSagzab(
            raw_data_path,
            socio_output_path,
            format_path,
            testing=True,
            combine_excels=args.run_combine_data,
            run_classification=run_classification,
        )
        run_obj.run()

    else:
        print(
            parser.error("Nothing to run. --run-socio or --run-sagabz must be choosed")
        )
