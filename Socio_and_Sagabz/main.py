#!/usr/bin/python
import argparse
from run_sagabz_socio import RunSagzab, RunSocio
import os

FORMAT_PATH = os.getcwd() + r"\Formats\socio_format_new.docx"

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument('--run-socio', action='store_true', default=False, help="run the socio option")
    parser.add_argument('--run-sagabz', action='store_true', default=False, help="run the sagabz option")
    parser.add_argument('--dont-run-classification', action='store_true', default=False, help="when used, ai classifications will not execute.")
    
    parser.add_argument('--start-task', choices=['split_excel', 'statistics', 'word_build'], default=None,
                        help="subtask to begin with.\
                              when we use this option, sociometry_output/word/combined_data.xlsx MUST exist.\
                              split_excel: split the combined_data.xlsx to individual excel files.\
                              statistics: run statistics on the individual excel files.\
                              word_build: build the word files.\
                              relevant for --run-socio only")
    parser.add_argument('--start-cadet', type=str, default=None,
                        help="when using --start-point split_excel/word_build, we might want to start\
                              we might want to start with a specific cadet due to a crash or manual change during run.\
                              relevant for --run-socio only")

    parser.add_argument('--socio-output-path', type=str, default=r"sociometry_output", help="path for socio ouptut")
    parser.add_argument('--sagabz-output-path', type=str, default=r"sagaz_output", help="path for sagabz output")
    parser.add_argument('--format-path', type=str, default="Formats/socio_format_for_classification.docx", help='path for the format file.')
    parser.add_argument('--old-stats-path', type=str, default=None, help="old stats path.")
    parser.add_argument('--raw-data-path', type=str, default=r"Excels", help="path for excels files.")
    parser.add_argument('--names-to-hashes', type=str, default=r"False", help="convert names to hashes. for unanimous data.")

    args = parser.parse_args()
    
    relative_path = os.getcwd()
    
    socio_output_path = os.path.join(relative_path, args.socio_output_path)
    sagabz_output_path = os.path.join(relative_path, args.sagabz_output_path)
    format_path = FORMAT_PATH
    run_classification = not args.dont_run_classification
    start_task = args.start_task
    start_cadet = args.start_cadet
    
    if args.old_stats_path is not None:
        old_stats_path = os.path.join(relative_path, args.old_stats_path)
    else:
        old_stats_path = None
    raw_data_path = os.path.join(relative_path, args.raw_data_path)

    if args.run_socio:
        run_obj = RunSocio(raw_data_path,
                           socio_output_path,
                           format_path,
                           old_stats_path=old_stats_path,
                           testing=True,
                           run_classification=run_classification,
                           start_task=start_task,
                           start_cadet=start_cadet,
                           names_to_hashes=args.names_to_hashes)
        run_obj.run()

    elif args.run_sagabz:
        run_obj = RunSagzab(raw_data_path,
                           socio_output_path,
                           format_path,
                           old_stats_path=old_stats_path,
                           testing=True,
                           run_classification=run_classification,
                           start_task=start_task,
                           start_cadet=start_cadet)
        run_obj.run()

    else:
        print(parser.error("Nothing to run. --run-socio or --run-sagabz must be choosed"))
